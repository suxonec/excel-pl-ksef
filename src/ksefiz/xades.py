"""XAdES-BES enveloped signature of an AuthTokenRequest, plus a self-signed company-seal
certificate for the KSeF TEST environment (self-signed certificates are rejected on DEMO/PROD).

Signature layout follows the MF C# client (KSeF.Client/Api/Services/SignatureService.cs):
RSA-SHA256, exclusive C14N, a reference to the whole document (enveloped) and one to
xades:SignedProperties, KeyInfo with the certificate.
"""

from __future__ import annotations

import base64
import hashlib
from datetime import UTC, datetime, timedelta

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa
from cryptography.x509.oid import NameOID
from lxml import etree  # type: ignore[import-untyped]

AUTH_NS = "http://ksef.mf.gov.pl/auth/token/2.0"
DS = "http://www.w3.org/2000/09/xmldsig#"
XADES = "http://uri.etsi.org/01903/v1.3.2#"
EXC_C14N = "http://www.w3.org/2001/10/xml-exc-c14n#"
ENVELOPED = "http://www.w3.org/2000/09/xmldsig#enveloped-signature"
RSA_SHA256 = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"
SHA256 = "http://www.w3.org/2001/04/xmlenc#sha256"
SIGNED_PROPS_TYPE = "http://uri.etsi.org/01903#SignedProperties"

ORGANIZATION_IDENTIFIER = x509.ObjectIdentifier("2.5.4.97")


def self_signed_seal(nip: str, name: str = "ksefiz TEST seal") -> tuple[x509.Certificate,
                                                                         rsa.RSAPrivateKey]:
    """Company-seal equivalent: organizationIdentifier VATPL-<nip>. TEST environment only."""
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = x509.Name([
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, name),
        x509.NameAttribute(ORGANIZATION_IDENTIFIER, f"VATPL-{nip}"),
        x509.NameAttribute(NameOID.COMMON_NAME, name),
        x509.NameAttribute(NameOID.COUNTRY_NAME, "PL"),
    ])
    now = datetime.now(UTC)
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(subject)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(minutes=61))
        .not_valid_after(now + timedelta(days=730))
        .sign(key, hashes.SHA256())
    )
    return cert, key


def save_pem(cert: x509.Certificate, key: rsa.RSAPrivateKey) -> bytes:
    return key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8,
                             serialization.NoEncryption()) + cert.public_bytes(
        serialization.Encoding.PEM)


def load_pem(data: bytes) -> tuple[x509.Certificate, rsa.RSAPrivateKey]:
    key = serialization.load_pem_private_key(data, password=None)
    if not isinstance(key, rsa.RSAPrivateKey):
        raise ValueError("expected an RSA private key")
    return x509.load_pem_x509_certificate(data), key


def auth_token_request(challenge: str, nip: str) -> etree._Element:
    root = etree.Element(f"{{{AUTH_NS}}}AuthTokenRequest", nsmap={None: AUTH_NS})
    etree.SubElement(root, f"{{{AUTH_NS}}}Challenge").text = challenge
    ctx = etree.SubElement(root, f"{{{AUTH_NS}}}ContextIdentifier")
    etree.SubElement(ctx, f"{{{AUTH_NS}}}Nip").text = nip
    etree.SubElement(root, f"{{{AUTH_NS}}}SubjectIdentifierType").text = "certificateSubject"
    return root


def _c14n(el: etree._Element) -> bytes:
    out: bytes = etree.tostring(el, method="c14n", exclusive=True, with_comments=False)
    return out


def _b64sha256(data: bytes) -> str:
    return base64.b64encode(hashlib.sha256(data).digest()).decode("ascii")


def _ds(parent: etree._Element, tag: str, **attrib: str) -> etree._Element:
    return etree.SubElement(parent, f"{{{DS}}}{tag}", attrib)


def _reference(signed_info: etree._Element, uri: str, digest: str, transforms: list[str],
               ref_type: str | None = None) -> None:
    attrs = {"URI": uri}
    if ref_type:
        attrs["Type"] = ref_type
    ref = _ds(signed_info, "Reference", **attrs)
    tr = _ds(ref, "Transforms")
    for t in transforms:
        _ds(tr, "Transform", Algorithm=t)
    _ds(ref, "DigestMethod", Algorithm=SHA256)
    _ds(ref, "DigestValue").text = digest


def sign(root: etree._Element, cert: x509.Certificate, key: rsa.RSAPrivateKey,
         signing_time: datetime | None = None) -> bytes:
    """Append an enveloped XAdES-BES signature to `root`; returns the serialised document."""
    doc_digest = _b64sha256(_c14n(root))  # enveloped transform: digest before the signature

    sig = etree.SubElement(root, f"{{{DS}}}Signature", {"Id": "Signature"}, nsmap={"ds": DS})
    signed_info = _ds(sig, "SignedInfo")
    _ds(signed_info, "CanonicalizationMethod", Algorithm=EXC_C14N)
    _ds(signed_info, "SignatureMethod", Algorithm=RSA_SHA256)
    sig_value = _ds(sig, "SignatureValue")
    key_info = _ds(sig, "KeyInfo")
    _ds(_ds(key_info, "X509Data"), "X509Certificate").text = base64.b64encode(
        cert.public_bytes(serialization.Encoding.DER)).decode("ascii")

    obj = _ds(sig, "Object")
    qp = etree.SubElement(obj, f"{{{XADES}}}QualifyingProperties", {"Target": "#Signature"},
                          nsmap={"xades": XADES})
    sp = etree.SubElement(qp, f"{{{XADES}}}SignedProperties", {"Id": "SignedProperties"})
    ssp = etree.SubElement(sp, f"{{{XADES}}}SignedSignatureProperties")
    when = (signing_time or datetime.now(UTC)) - timedelta(minutes=1)
    etree.SubElement(ssp, f"{{{XADES}}}SigningTime").text = when.isoformat(timespec="seconds")
    c = etree.SubElement(etree.SubElement(ssp, f"{{{XADES}}}SigningCertificate"),
                         f"{{{XADES}}}Cert")
    cd = etree.SubElement(c, f"{{{XADES}}}CertDigest")
    _ds(cd, "DigestMethod", Algorithm=SHA256)
    _ds(cd, "DigestValue").text = _b64sha256(cert.public_bytes(serialization.Encoding.DER))
    iss = etree.SubElement(c, f"{{{XADES}}}IssuerSerial")
    _ds(iss, "X509IssuerName").text = cert.issuer.rfc4514_string()
    _ds(iss, "X509SerialNumber").text = str(cert.serial_number)

    _reference(signed_info, "", doc_digest, [ENVELOPED, EXC_C14N])
    _reference(signed_info, "#SignedProperties", _b64sha256(_c14n(sp)), [EXC_C14N],
               SIGNED_PROPS_TYPE)

    signature = key.sign(_c14n(signed_info), padding.PKCS1v15(), hashes.SHA256())
    sig_value.text = base64.b64encode(signature).decode("ascii")
    out: bytes = etree.tostring(root, xml_declaration=True, encoding="utf-8")
    return out
