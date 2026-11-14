"""Consulta de contribuyentes en el padrón de ARCA (ws_sr_constancia_inscripcion).

Autenticación WSAA con certificado propio (firma CMS con `cryptography`, sin openssl)
y consulta SOAP `getPersona_v2`. Sin dependencias de Streamlit: las credenciales se
pasan como texto PEM.
"""
import base64
import hashlib
import json
import os
import re
import tempfile
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from html import escape

import requests
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.serialization import pkcs7

# ─── Producción ────────────────────────────────────────────────────────────────
WSAA_URL = "https://wsaa.afip.gov.ar/ws/services/LoginCms"
PADRON_URL = "https://aws.afip.gov.ar/sr-padron/webservices/personaServiceA5"
SERVICIO = "ws_sr_constancia_inscripcion"

TZ_AR = timezone(timedelta(hours=-3))
_TIMEOUT = 30
_MARGEN_TA = 300  # segundos antes del vencimiento en que el TA se considera vencido

_cache_ta = {}  # huella del certificado -> {token, sign, expiracion (epoch)}


class ArcaError(Exception):
    """Error devuelto por ARCA (SOAP Fault o respuesta sin datos)."""


# ─── CUIT ──────────────────────────────────────────────────────────────────────
def validar_cuit(cuit):
    """Devuelve el CUIT con 11 dígitos o lanza ValueError si no es válido."""
    c = re.sub(r"\D", "", str(cuit or ""))
    if len(c) != 11:
        raise ValueError("El CUIT debe tener 11 dígitos.")
    pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    resto = sum(int(d) * p for d, p in zip(c[:10], pesos)) % 11
    dv = 0 if resto == 0 else (9 if resto == 1 else 11 - resto)
    if dv != int(c[10]):
        raise ValueError("El dígito verificador del CUIT no es válido.")
    return c


def formatear_cuit(cuit):
    c = re.sub(r"\D", "", str(cuit))
    return f"{c[:2]}-{c[2:10]}-{c[10:]}" if len(c) == 11 else c


# ─── Utilidades XML ────────────────────────────────────────────────────────────
def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _hijo(el, nombre):
    if el is None:
        return None
    for h in el:
        if _local(h.tag) == nombre:
            return h
    return None


def _hijos(el, nombre):
    return [h for h in el if _local(h.tag) == nombre] if el is not None else []


def _txt(el, nombre, default=""):
    h = _hijo(el, nombre)
    return (h.text or "").strip() if h is not None and h.text else default


def _buscar(el, nombre):
    for x in el.iter():
        if _local(x.tag) == nombre:
            return x
    return None


def _post_soap(url, cuerpo):
    try:
        resp = requests.post(
            url, data=cuerpo.encode("utf-8"), timeout=_TIMEOUT,
            headers={"Content-Type": "text/xml; charset=utf-8", "SOAPAction": '""'},
        )
    except requests.RequestException as e:
        raise ArcaError(f"No se pudo conectar con ARCA: {e}") from e
    try:
        raiz = ET.fromstring(resp.content)
    except ET.ParseError as e:
        raise ArcaError(f"Respuesta inválida de ARCA (HTTP {resp.status_code}).") from e
    fault = _buscar(raiz, "Fault")
    if fault is not None:
        msg = _txt(fault, "faultstring") or "Error desconocido"
        raise ArcaError(msg)
    if resp.status_code >= 400:
        raise ArcaError(f"ARCA respondió HTTP {resp.status_code}.")
    return raiz


# ─── WSAA ──────────────────────────────────────────────────────────────────────
def _crear_tra(servicio=SERVICIO):
    ahora = datetime.now(TZ_AR).replace(microsecond=0)
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<loginTicketRequest version="1.0"><header>'
        f"<uniqueId>{int(time.time())}</uniqueId>"
        f"<generationTime>{(ahora - timedelta(minutes=10)).isoformat()}</generationTime>"
        f"<expirationTime>{(ahora + timedelta(minutes=10)).isoformat()}</expirationTime>"
        f"</header><service>{servicio}</service></loginTicketRequest>"
    )


def _firmar_cms(tra, cert_pem, key_pem):
    """Firma el TRA como CMS con el contenido incluido (DER, base64)."""
    cert = x509.load_pem_x509_certificate(cert_pem.encode() if isinstance(cert_pem, str) else cert_pem)
    key = serialization.load_pem_private_key(
        key_pem.encode() if isinstance(key_pem, str) else key_pem, password=None
    )
    der = (
        pkcs7.PKCS7SignatureBuilder()
        .set_data(tra.encode("utf-8"))
        .add_signer(cert, key, hashes.SHA256())
        .sign(serialization.Encoding.DER, [])
    )
    return base64.b64encode(der).decode("ascii")


def _huella(cert_pem):
    return hashlib.sha256(cert_pem.strip().encode()).hexdigest()[:16]


def _archivo_cache(huella):
    return os.path.join(tempfile.gettempdir(), f"arca_ta_{SERVICIO}_{huella}.json")


def _ta_vigente(ta):
    return bool(ta) and ta.get("expiracion", 0) - _MARGEN_TA > time.time()


def obtener_ticket(cert_pem, key_pem):
    """Devuelve {token, sign, expiracion}; reutiliza el TA cacheado mientras esté vigente
    (WSAA rechaza pedir otro si ya hay uno válido para el mismo certificado)."""
    huella = _huella(cert_pem)
    ta = _cache_ta.get(huella)
    if _ta_vigente(ta):
        return ta
    try:
        with open(_archivo_cache(huella), encoding="utf-8") as f:
            ta = json.load(f)
        if _ta_vigente(ta):
            _cache_ta[huella] = ta
            return ta
    except (OSError, ValueError):
        pass

    cms = _firmar_cms(_crear_tra(), cert_pem, key_pem)
    cuerpo = (
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:wsaa="http://wsaa.view.sua.dvadac.desein.afip.gov">'
        "<soapenv:Header/><soapenv:Body><wsaa:loginCms>"
        f"<wsaa:in0>{cms}</wsaa:in0>"
        "</wsaa:loginCms></soapenv:Body></soapenv:Envelope>"
    )
    try:
        raiz = _post_soap(WSAA_URL, cuerpo)
    except ArcaError as e:
        if "ya posee un TA valido" in str(e) or "ya posee un TA válido" in str(e):
            raise ArcaError(
                "ARCA ya entregó un ticket vigente para este certificado y la app no lo tiene "
                "guardado (probablemente se reinició). Hay que esperar a que venza (hasta 12 h) "
                "o usar otro certificado."
            ) from e
        raise
    ret = _buscar(raiz, "loginCmsReturn")
    if ret is None or not (ret.text or "").strip():
        raise ArcaError("WSAA no devolvió el ticket de acceso.")
    ticket = ET.fromstring(ret.text.strip().encode("utf-8"))
    exp_txt = _txt(_buscar(ticket, "header"), "expirationTime")
    ta = {
        "token": _txt(_buscar(ticket, "credentials"), "token"),
        "sign": _txt(_buscar(ticket, "credentials"), "sign"),
        "expiracion": datetime.fromisoformat(exp_txt).timestamp() if exp_txt else time.time() + 11 * 3600,
    }
    if not ta["token"] or not ta["sign"]:
        raise ArcaError("WSAA devolvió un ticket incompleto.")
    _cache_ta[huella] = ta
    try:
        with open(_archivo_cache(huella), "w", encoding="utf-8") as f:
            json.dump(ta, f)
    except OSError:
        pass
    return ta


# ─── Padrón: getPersona_v2 ─────────────────────────────────────────────────────
def _errores(el):
    if el is None:
        return []
    errs = [(e.text or "").strip() for e in _hijos(el, "error") if (e.text or "").strip()]
    msg = _txt(el, "mensaje")
    return errs + ([msg] if msg else [])


def parsear_persona(raiz):
    """Normaliza la respuesta de getPersona_v2 a un dict."""
    pr = _buscar(raiz, "personaReturn")
    if pr is None:
        raise ArcaError("ARCA no devolvió datos para ese CUIT.")
    dg = _hijo(pr, "datosGenerales")
    drg = _hijo(pr, "datosRegimenGeneral")
    dm = _hijo(pr, "datosMonotributo")

    errores = []
    for nombre in ("errorConstancia", "errorRegimenGeneral", "errorMonotributo"):
        errores += _errores(_hijo(pr, nombre))

    ec = _hijo(pr, "errorConstancia")
    base = dg if dg is not None else ec
    razon = _txt(base, "razonSocial")
    nombre = razon or " ".join(x for x in (_txt(base, "apellido"), _txt(base, "nombre")) if x)

    dom_el = _hijo(dg, "domicilioFiscal")
    domicilio = {
        "direccion": _txt(dom_el, "direccion"),
        "localidad": _txt(dom_el, "localidad"),
        "cod_postal": _txt(dom_el, "codPostal"),
        "provincia": _txt(dom_el, "descripcionProvincia"),
    }

    def _acts(padre, etiqueta):
        out = []
        for a in _hijos(padre, etiqueta):
            out.append({
                "id": _txt(a, "idActividad"),
                "descripcion": _txt(a, "descripcionActividad"),
                "orden": int(_txt(a, "orden", "0") or 0),
                "periodo": _txt(a, "periodo"),
            })
        return out

    actividades = sorted(_acts(drg, "actividad"), key=lambda a: (a["orden"] == 0, a["orden"]))
    impuestos = []
    for padre in (drg, dm):
        for i in _hijos(padre, "impuesto"):
            impuestos.append({
                "id": _txt(i, "idImpuesto"),
                "descripcion": _txt(i, "descripcionImpuesto"),
                "estado": _txt(i, "estadoImpuesto"),
                "periodo": _txt(i, "periodo"),
            })

    monotributo = None
    if dm is not None:
        cat = _hijo(dm, "categoriaMonotributo")
        act_m = _hijo(dm, "actividadMonotributista")
        monotributo = {
            "categoria": _txt(cat, "descripcionCategoria"),
            "actividad": _txt(act_m, "descripcionActividad"),
            "id_actividad": _txt(act_m, "idActividad"),
        }
        if not actividades and act_m is not None:
            actividades = [{
                "id": monotributo["id_actividad"], "descripcion": monotributo["actividad"],
                "orden": 1, "periodo": _txt(act_m, "periodo"),
            }]

    ids_imp = {i["id"] for i in impuestos if i["estado"] in ("", "ACTIVO", "AC")}
    if monotributo is not None or "20" in ids_imp:
        condicion_iva = "Monotributo"
    elif "30" in ids_imp:
        condicion_iva = "IVA Responsable Inscripto"
    elif "32" in ids_imp:
        condicion_iva = "IVA Exento"
    else:
        condicion_iva = "No inscripto / no informado"

    return {
        "cuit": _txt(base, "idPersona"),
        "nombre": nombre,
        "tipo_persona": _txt(dg, "tipoPersona"),
        "estado_clave": _txt(dg, "estadoClave"),
        "domicilio": domicilio,
        "actividades": actividades,
        "impuestos": impuestos,
        "monotributo": monotributo,
        "condicion_iva": condicion_iva,
        "errores": errores,
    }


def consultar_cuit(cuit, cert_pem, key_pem, cuit_representada):
    """Consulta un CUIT en el padrón de ARCA y devuelve el dict de `parsear_persona`."""
    cuit = validar_cuit(cuit)
    representada = re.sub(r"\D", "", str(cuit_representada))
    ta = obtener_ticket(cert_pem, key_pem)
    cuerpo = (
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:a5="http://a5.soap.ws.server.puc.sr/">'
        "<soapenv:Header/><soapenv:Body><a5:getPersona_v2>"
        f"<token>{escape(ta['token'])}</token>"
        f"<sign>{escape(ta['sign'])}</sign>"
        f"<cuitRepresentada>{representada}</cuitRepresentada>"
        f"<idPersona>{cuit}</idPersona>"
        "</a5:getPersona_v2></soapenv:Body></soapenv:Envelope>"
    )
    datos = parsear_persona(_post_soap(PADRON_URL, cuerpo))
    if not datos["cuit"]:
        datos["cuit"] = cuit
    return datos
