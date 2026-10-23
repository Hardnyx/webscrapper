"""Standalone page-function check for SBS curves; no repository installation needed."""
import argparse
from importlib.metadata import PackageNotFoundError, version
import json
from pathlib import Path
import ssl
import subprocess
import sys


def dependencies():
    try:
        if not 24 <= int(version('packaging').split('.')[0]) < 27:
            raise PackageNotFoundError('packaging')
    except PackageNotFoundError:
        subprocess.check_call([sys.executable,'-m','pip','install','packaging>=24,<27'])
    import importlib, site
    if site.ENABLE_USER_SITE:site.addsitedir(site.getusersitepackages())
    importlib.invalidate_caches()
    from packaging.specifiers import SpecifierSet
    for name,bounds in [('curl_cffi','>=0.16,<1'),('certifi','>=2024.2.2')]:
        try:compatible=version(name) in SpecifierSet(bounds)
        except PackageNotFoundError:compatible=False
        if not compatible:subprocess.check_call([sys.executable,'-m','pip','install',name+bounds])
        print('[deps]',name,version(name))
    if site.ENABLE_USER_SITE:site.addsitedir(site.getusersitepackages())
    importlib.invalidate_caches()


def main():
    from datetime import date
    p=argparse.ArgumentParser(description='Prueba del POST JSON de curvas SBS con TLS verificado.')
    p.add_argument('--fecha',required=True);p.add_argument('--curva',default='CCPSS')
    p.add_argument('--salida',type=Path,default=Path.home()/'fuentes_financieras_diagnosticos'/'curva_sbs')
    a=p.parse_args();date.fromisoformat(a.fecha)
    if a.curva not in {'CCPSS','CBCRPS','CCPVS','CCPEDS','CBCRS','CCSDF','CSBCRD','CCINFS','CCCLD'}:p.error('Curva no reconocida.')
    dependencies()
    import certifi
    from curl_cffi import requests
    folder=a.salida.resolve();folder.mkdir(parents=True,exist_ok=True)
    bundle=Path(certifi.where()).read_text(encoding='ascii')
    for der in ssl.create_default_context().get_ca_certs(binary_form=True):bundle+='\n'+ssl.DER_cert_to_PEM_cert(der)
    ca=folder/'ca_verificada.pem';ca.write_text(bundle,encoding='ascii')
    url='https://www.sbs.gob.pe/app/pp/n_CurvaSoberana/BuscarInformacionHistorica'
    result={'fecha':a.fecha,'curva':a.curva,'url':url,'estado':'No ejecutado'}
    try:
        with requests.Session(impersonate='chrome') as session:
            r=session.post(url,json={'tipoCurva':a.curva,'fechaInicio':a.fecha,'fechaFin':a.fecha},
                headers={'Referer':'https://www.sbs.gob.pe/app/pp/n_CurvaSoberana/CurvaSoberana/ConsultaHistorica'},
                verify=str(ca),timeout=120)
        result.update(http=r.status_code,bytes=len(r.content));(folder/'respuesta.txt').write_bytes(r.content)
        low=r.text.lower()
        if r.status_code in (403,406,429,503) or any(x in low for x in ['solicitud rechazada','request blocked','access denied','incapsula','políticas de seguridad']):
            result['estado']='Bloqueo HTTP/WAF'
        elif r.status_code>=400:result['estado']='Error HTTP'
        else:
            try:body=json.loads(r.content.decode('utf-8-sig'))
            except (ValueError,UnicodeError):body=None
            if not isinstance(body,dict) or body.get('success') is not True or not isinstance(body.get('data'),list):result['estado']='Error de estructura: JSON no válido para curvas'
            elif not body['data']:result['estado']='Consulta válida sin puntos; no acredita captura financiera'
            elif any(str(row.get('fechaProceso',''))[:10]!=a.fecha or any(k not in row for k in ('plazos','tasa')) for row in body['data']):result['estado']='Error de estructura o fecha distinta'
            else:result.update(estado='Página funcional: puntos de la fecha solicitada',filas=len(body['data']))
    except Exception as exc:
        result.update(estado='Fallo de conexión o TLS',detalle=type(exc).__name__)
    (folder/'diagnostico.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False,indent=2))
    return 0 if result['estado'].startswith('Página funcional') else 2

if __name__=='__main__':raise SystemExit(main())
