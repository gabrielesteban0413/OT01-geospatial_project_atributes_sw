import re
import pandas as pd
from pathlib import Path
from typing import Optional, Tuple, List, Dict, Any
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

RUTA_ORIGEN = Path(r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\02_ConversionCAD_Shape\Archivos_Service_PATH\Path 28 Sept")
RUTA_DESTINO = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output")
RUTA_CENTRALES = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\03-props_centrales.xlsx")
NOMBRE_SALIDA = "01-Bulk_Path.xlsx"
EXTENSION = ".txt"
VACIO = "-"

HOJA_EC = 'EC'
COLUMNA_EC = 'cabecera'
HOJA_BLACKLIST = 'BLACKLIST'
COLUMNA_BLACKLIST = 'BLACKLIST'
COLUMNA_DESTINO = 'COLUMNA'


def _c(pats: List[str]) -> List[re.Pattern]:
    return [re.compile(p, re.IGNORECASE | re.MULTILINE) for p in pats]


_INTF = (
    r'(?:'
    r'TwentyFiveGigE|FortyGigabitEthernet|TenGigabitEthernet|GigabitEthernet|FastEthernet|Ethernet'
    r'|TW|FO|TE|Te|GE|Gi|FE|Fa|ETH|Et'
    r')'
)

P = {
    'ip_demarcador': _c([
        r'IP\s+DEMARCADOR\s*:?\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'IP\s+DEMARCADOR\s*:?\s*\n\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'DEMARCADOR\s*:?\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'DEMARCADOR\s*:?\s*\n\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'IP\s*DEMAR\s*:?\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'IP\s*DEMAR\s*:?\s*\n\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
    ]),
    'ip_ping': _c([
        r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})[^\n]*?\bping\b',
        r'\bping\b[^\n]*?(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
    ]),
    'anillo_origen': _c([
        r'ANILLO\s+ORIGEN\s*:\s*([A-Z0-9/\-]{2,20})',
        r'ANILLO\s+ORIGEN\s+([A-Z0-9/\-]{2,20})',
    ]),
    'anillo_destino': _c([
        r'ANILLO\s+DESTINO\s*:\s*([A-Z0-9/\-]{2,20})',
        r'ANILLO\s+DESTINO\s+([A-Z0-9/\-]{2,20})',
    ]),
    'anillo': _c([
        r'ANILLO\s*O\s*NODO\s*:\s*([A-Z0-9/\-]{6,20})',
        r'ANILLO\s*:\s*([A-Z0-9/\-]{6,20})',
        r'RING\s*:\s*([A-Z0-9/\-]{6,20})',
        r'SCJ\s*:\s*([A-Z0-9/\-]{6,20})',
        r'ORIGEN\s*:\s*([A-Z0-9/\-]{6,20})',
    ]),
    'ancho_banda': _c([
        r'ANCHO DE BANDA FINAL\s*[:=]?\s*([^\n]+)',
        r'BW:\s*([^\n]+)',
        r'BW\s+([^\n]+)',
        r'ANCHO DE BANDA SOLICITADO\s*([^\n]+)',
    ]),
    'id_servicio': _c([
        r'ID\s+DEL\s+SERVICIO\s+AFECTADO\s*:\s*([A-Z0-9]+)',
        r'ID\s+DEL\s+SERVICIO\s+AFECTADO\s*:\s*([^\n]+)',
        r'ID\s+SERVICIO\s*:\s*([A-Z0-9]+)',
        r'ID\s+SERVICIO\s*:\s*([^\n]+)',
        r'ID\s+DEL\s+SERVICIO\s+AFECTADO\s+([A-Z0-9]+)',
        r'ID\s+SERVICIO\s+([A-Z0-9]+)',
        r'\b(CAV\d{2}[A-Z]{2}\d{7,10})\b',
        r'\b([A-Z]{3}\d{2}[A-Z]{2}\d{7,10})\b',
    ]),
    'coord_multipoint': _c([
        r'COORDENADAS:\s*MultiPointZ\s*\(\s*\(\s*([-+]?\d+\.\d+)\s+([-+]?\d+\.\d+)',
        r'MultiPointZ\s*\(\s*\(\s*([-+]?\d+\.\d+)\s+([-+]?\d+\.\d+)',
    ]),
    'lat_decimal': _c([r'LATITUD:\s*([-+]?\s*\d+\.\d+)']),
    'lon_decimal': _c([r'LONGITUD:\s*([-+]?\s*\d+\.\d+)']),
    'lat_gms':     _c([r'LATITUD:\s*([0-9°\'\"\.]+)']),
    'lon_gms':     _c([r'LONGITUD:\s*([0-9°\'\"\.]+)']),
    'tipo_serv':   _c([
        r'TIPO DE SERVICIO:\s*([^\n]+)',
        r'TIPO DE SERVICIO\s+([^\n]+)',
    ]),
    'puerto_common': _c([
        r'PUERTO COMMON:\s*([^\n]+)',
        r'PUERTO COMMON\s+([^\n]+)',
    ]),
    'puerto_owner': _c([
        r'PUERTO OWNER:\s*([^\n]+)',
        r'PUERTO OWNER\s+([^\n]+)',
    ]),
    'recursos_metro': re.compile(r'RECURSOS METRO.*?<([A-Z0-9]+)>', re.IGNORECASE | re.DOTALL),
    'interfaz': re.compile(
        rf'\b{_INTF}\s*[1-9]\s*/\s*\d+(?:\s*/\s*\d+)*',
        re.IGNORECASE,
    ),
    'common_intf': re.compile(
        rf'(?:COMMON|Common)[- ]*{_INTF}\s*([1-9]\s*/\s*\d+(?:\s*/\s*\d+)*)',
        re.IGNORECASE,
    ),
    'owner_intf': re.compile(
        rf'RPL\s+OWNER[- ]*{_INTF}\s*([1-9]\s*/\s*\d+(?:\s*/\s*\d+)*)',
        re.IGNORECASE,
    ),
    'intf_common': re.compile(
        rf'({_INTF}\s*[1-9]\s*/\s*\d+(?:\s*/\s*\d+)*)\s+Common',
        re.IGNORECASE,
    ),
    'intf_owner': re.compile(
        rf'({_INTF}\s*[1-9]\s*/\s*\d+(?:\s*/\s*\d+)*)\s+RPL\s+Owner',
        re.IGNORECASE,
    ),
    'interfaz_corchete': re.compile(
        rf'\[\s*[A-Z0-9]+[-\s]*({_INTF}\s*[1-9]\s*/\s*\d+(?:\s*/\s*\d+)*)',
        re.IGNORECASE,
    ),
    'cambio': re.compile(
        r'\bOC-\d+'
        r'|CAMBIO(?:\s+(?:APROVISIONAMIENTO|DE\s+FACTIBILIDAD))?\s*:?\s*(C\d{4,10})\b',
        re.IGNORECASE,
    ),
    'etiqueta_todem': re.compile(r'TO_DEM_[^\s]*', re.IGNORECASE),
    'etiqueta_todem_eol': re.compile(r'TO_DEM_[^\n]*', re.IGNORECASE),
    'equipo_generico': re.compile(r'\b([A-Z]{2,}[A-Z0-9]{4,})\b'),
    'gms_parse': re.compile(r'(\d{1,3})[°]?(\d{2})\'(\d{1,3}(?:\.\d+)?)'),
    'bw_numero': re.compile(r'(\d+(?:\.\d+)?)'),
}

LINEAS_CLAVE_EC = ('METRO:', 'CABECERA:', 'PRINCIPAL:', 'ACCESO:')
LINEAS_ETIQUETA = ('ETIQUETA METRO:', 'DESCRIPCION:', 'DESCRIPTION:')

PREFIJOS_ETHERNET = [
    'TwentyFiveGigE', 'FortyGigabitEthernet', 'TenGigabitEthernet',
    'GigabitEthernet', 'FastEthernet', 'Ethernet',
    'TW', 'FO', 'TE', 'GE', 'GI', 'FE', 'FA', 'ETH',
]


class CoordenadasFormatter:
    @staticmethod
    def a_gms(valor: float) -> str:
        neg = valor < 0
        v = abs(valor)
        g = int(v)
        md = (v - g) * 60
        m = int(md)
        s = (md - m) * 60
        signo = '-' if neg else ''
        return f"{signo}{g:02d}°{m:02d}'{s:.1f}\""

    @staticmethod
    def gms_str(coord: str) -> str:
        if not coord or coord in ('NO', 'NO ENCONTRADA', 'ERROR'):
            return coord
        m = P['gms_parse'].search(coord.replace('"', '').strip())
        if m:
            g = m.group(1).zfill(2) if len(m.group(1)) < 3 else m.group(1).zfill(3)
            return f"{g}°{m.group(2)}'{m.group(3)}\""
        return coord


class DataExtractor:
    def __init__(self):
        self.fmt = CoordenadasFormatter()
        self.equipos_validos: set = set()
        self.blacklist_global: set = set()
        self.blacklist_columna: Dict[str, set] = {}
        self._cargar_centrales()
        self._cargar_blacklist()

    def _cargar_centrales(self):
        try:
            if RUTA_CENTRALES.exists():
                df = pd.read_excel(
                    RUTA_CENTRALES, sheet_name=HOJA_EC,
                    engine='openpyxl', usecols=[COLUMNA_EC],
                )
                self.equipos_validos = set(
                    df[COLUMNA_EC].dropna().astype(str).str.strip()
                )
        except Exception:
            self.equipos_validos = set()

    def _cargar_blacklist(self):
        self.blacklist_global = {'NULL', 'NA', 'NAN', 'TODAS'}
        self.blacklist_columna = {
            'ANILLO': {
                'ANILLO', 'ANILLOS', 'RING', 'SCJ', 'ORIGEN', 'DESTINO',
                'NEMONICO', 'POR', 'NR', 'XXXXXXXXXXXXXXXXXXXX',
                'REALIZADO', 'PENDIENTE', 'ANULADO', 'CANCELADO',
                'EJECUTADO', 'ACTIVO', 'INACTIVO', 'NINGUNO', 'NINGUNA',
                'SINANILLO', 'TERMINADO', 'FINALIZADO', 'PROCESO',
                'PROCESADO', 'ABIERTO', 'CERRADO', 'ENPROCESO', 'NODEFINIDO',
            },
        }

        try:
            if not RUTA_CENTRALES.exists():
                return
            df = pd.read_excel(
                RUTA_CENTRALES, sheet_name=HOJA_BLACKLIST,
                engine='openpyxl',
            )
        except Exception:
            return

        cols = {c.lower(): c for c in df.columns}
        col_palabra = cols.get(COLUMNA_BLACKLIST.lower())
        col_columna = cols.get(COLUMNA_DESTINO.lower())
        if col_palabra is None:
            return

        for _, row in df.iterrows():
            palabra_raw = row[col_palabra]
            if pd.isna(palabra_raw):
                continue
            palabra = re.sub(r'[^A-Z0-9]', '', str(palabra_raw).upper())
            if not palabra:
                continue

            columna = ''
            if col_columna is not None and not pd.isna(row[col_columna]):
                columna = str(row[col_columna]).strip().upper()

            if not columna or columna == 'TODAS':
                self.blacklist_global.add(palabra)
            else:
                self.blacklist_columna.setdefault(columna, set()).add(palabra)

    def _es_valido(self, valor: str, columna: str) -> bool:
        if not valor:
            return False
        limpio = re.sub(r'[^A-Z0-9]', '', str(valor).upper())
        if not limpio:
            return False
        if limpio in self.blacklist_global:
            return False
        if limpio in self.blacklist_columna.get(columna, set()):
            return False
        return True

    @staticmethod
    def _ip_valida(ip: str) -> bool:
        partes = ip.split('.')
        return len(partes) == 4 and all(
            p.isdigit() and 0 <= int(p) <= 255 for p in partes
        )

    @staticmethod
    def _token_linea(linea: str) -> Optional[str]:
        if ':' not in linea:
            return None
        resto = linea.split(':', 1)[1].strip()
        for t in re.split(r'[\s\t]+', resto):
            if t:
                return t
        return None

    @staticmethod
    def _normalizar_interfaz(raw: str) -> str:
        compacto = re.sub(r'\s+', '', raw)
        low = compacto.lower()
        for largo in PREFIJOS_ETHERNET:
            if low.startswith(largo.lower()):
                num = compacto[len(largo):]
                if num.startswith('0'):
                    return ""
                return num
        if compacto.startswith('0'):
            return ""
        return compacto

    @staticmethod
    def _puerto_en_linea(linea: str) -> Optional[str]:
        m = P['interfaz'].search(linea)
        if not m:
            return None
        pto = DataExtractor._normalizar_interfaz(m.group(0))
        return pto or None

    @staticmethod
    def _limpiar_puerto(valor: str) -> str:
        if not valor or valor == VACIO or valor.startswith("SIN "):
            return ""
        m = P['interfaz'].search(valor)
        if m:
            pto = DataExtractor._normalizar_interfaz(m.group(0))
            if pto:
                return pto
        m = re.search(r'([1-9]\s*/\s*\d+(?:\s*/\s*\d+)*)', valor)
        if m:
            return re.sub(r'\s+', '', m.group(1))
        return ""

    @staticmethod
    def _todas_las_interfaces(texto: str) -> set:
        vistas = set()
        for m in P['interfaz'].finditer(texto):
            pto = DataExtractor._normalizar_interfaz(m.group(0))
            if pto:
                vistas.add(pto)
        return vistas

    @classmethod
    def _interfaz_unica_en_texto(cls, texto: str) -> Optional[str]:
        vistas = cls._todas_las_interfaces(texto)
        if len(vistas) == 1:
            return next(iter(vistas))
        return None

    def ip_demarcador(self, t: str) -> str:
        for pat in P['ip_demarcador']:
            m = pat.search(t)
            if m:
                ip = m.group(1).strip()
                if self._ip_valida(ip) and self._es_valido(ip, 'IP_DEMARCADOR'):
                    return ip

        lineas = t.splitlines()
        for i, linea in enumerate(lineas):
            if 'PING' not in linea.upper():
                continue
            contexto = ' '.join(
                lineas[j].upper()
                for j in range(max(0, i - 1), min(len(lineas), i + 2))
            )
            if not any(k in contexto for k in ('DEMARCADOR', 'DEMAR', 'DEM')):
                continue
            for pat in P['ip_ping']:
                m = pat.search(linea)
                if m:
                    ip = m.group(1).strip()
                    if self._ip_valida(ip) and self._es_valido(ip, 'IP_DEMARCADOR'):
                        return ip

        return VACIO

    def anillo(self, t: str) -> str:
        valor_origen = None
        for pat in P['anillo_origen']:
            m = pat.search(t)
            if m:
                valor_origen = m.group(1).strip().upper()
                break

        if valor_origen:
            cand = re.sub(r'[^A-Z0-9]', '', valor_origen)
            if self._es_valido(cand, 'ANILLO'):
                return cand
            for pat in P['anillo_destino']:
                m = pat.search(t)
                if m:
                    cand_d = re.sub(r'[^A-Z0-9]', '', m.group(1).upper())
                    if self._es_valido(cand_d, 'ANILLO'):
                        return cand_d

        for pat in P['anillo_destino']:
            m = pat.search(t)
            if m:
                cand = re.sub(r'[^A-Z0-9]', '', m.group(1).upper())
                if self._es_valido(cand, 'ANILLO'):
                    return cand

        for pat in P['anillo']:
            m = pat.search(t)
            if m:
                cand = re.sub(r'[^A-Z0-9]', '', m.group(1).upper())
                if len(cand) >= 6 and self._es_valido(cand, 'ANILLO'):
                    return cand

        return VACIO

    def ancho_banda(self, t: str) -> str:
        for pat in P['ancho_banda']:
            m = pat.search(t)
            if m:
                v = m.group(1).strip().strip(':,- ')
                num = P['bw_numero'].search(v)
                if num:
                    v = num.group(1)
                else:
                    v = re.sub(r'(?i)(mbps|mbit|megabite?)', '', v).strip()
                if v and self._es_valido(v, 'ANCHO_BANDA'):
                    return v
        return VACIO

    def id_servicio(self, t: str) -> str:
        for pat in P['id_servicio']:
            for m in pat.finditer(t):
                v = re.sub(r'[^\w]', '', m.group(1).strip().strip(':,- '))
                if v and len(v) >= 10 and self._es_valido(v, 'ID_SERVICIO'):
                    return v
        return VACIO

    def tipo_servicio(self, t: str) -> str:
        for pat in P['tipo_serv']:
            m = pat.search(t)
            if m:
                v = m.group(1).strip().strip(':,- ')
                if v and self._es_valido(v, 'TIPO_SERVICIO'):
                    return v
        return VACIO

    def equipo_central(self, t: str) -> str:
        catalogo = self.equipos_validos

        m = P['recursos_metro'].search(t)
        if m:
            cand = m.group(1).strip()
            if self._es_valido(cand, 'EQUIPO_CENTRAL'):
                if not catalogo or cand in catalogo:
                    return cand

        for linea in t.splitlines():
            lu = linea.upper()
            if any(k in lu for k in LINEAS_CLAVE_EC):
                tok = self._token_linea(linea)
                if tok and self._es_valido(tok, 'EQUIPO_CENTRAL'):
                    if not catalogo or tok in catalogo:
                        return tok

        if catalogo:
            for cand in P['equipo_generico'].findall(t):
                if cand in catalogo and self._es_valido(cand, 'EQUIPO_CENTRAL'):
                    return cand

        if not catalogo:
            for cand in P['equipo_generico'].findall(t):
                if self._es_valido(cand, 'EQUIPO_CENTRAL'):
                    return cand

        return VACIO

    def puerto_common(self, t: str) -> str:
        for linea in t.splitlines():
            if 'COMMON' in linea.upper():
                pto = self._puerto_en_linea(linea)
                if pto and self._es_valido(pto, 'PUERTO_COMMON'):
                    return pto

        for pat in P['puerto_common']:
            m = pat.search(t)
            if m:
                v = self._limpiar_puerto(m.group(1).strip().strip(':,- '))
                if v and self._es_valido(v, 'PUERTO_COMMON'):
                    return v

        m = P['common_intf'].search(t)
        if m:
            v = self._limpiar_puerto(m.group(0))
            if v and self._es_valido(v, 'PUERTO_COMMON'):
                return v

        m = P['intf_common'].search(t)
        if m:
            v = self._limpiar_puerto(m.group(1))
            if v and self._es_valido(v, 'PUERTO_COMMON'):
                return v

        v = self._interfaz_unica_en_texto(t)
        if v and self._es_valido(v, 'PUERTO_COMMON'):
            return v
        return VACIO

    def puerto_owner(self, t: str) -> str:
        for linea in t.splitlines():
            if 'OWNER' in linea.upper():
                pto = self._puerto_en_linea(linea)
                if pto and self._es_valido(pto, 'PUERTO_OWNER'):
                    return pto

        for pat in P['puerto_owner']:
            m = pat.search(t)
            if m:
                v = self._limpiar_puerto(m.group(1).strip().strip(':,- '))
                if v and self._es_valido(v, 'PUERTO_OWNER'):
                    return v

        m = P['owner_intf'].search(t)
        if m:
            v = self._limpiar_puerto(m.group(0))
            if v and self._es_valido(v, 'PUERTO_OWNER'):
                return v

        m = P['intf_owner'].search(t)
        if m:
            v = self._limpiar_puerto(m.group(1))
            if v and self._es_valido(v, 'PUERTO_OWNER'):
                return v

        v = self._interfaz_unica_en_texto(t)
        if v and self._es_valido(v, 'PUERTO_OWNER'):
            return v
        return VACIO

    def coordenadas(self, t: str) -> Tuple[Optional[str], Optional[str]]:
        for pat in P['coord_multipoint']:
            m = pat.search(t)
            if m:
                lon = float(m.group(1))
                lat = float(m.group(2))
                if lon > 0:
                    lon = -lon
                return self.fmt.a_gms(lat), self.fmt.a_gms(lon)

        mlat = next((p.search(t) for p in P['lat_decimal'] if p.search(t)), None)
        mlon = next((p.search(t) for p in P['lon_decimal'] if p.search(t)), None)
        if mlat and mlon:
            try:
                lat = float(mlat.group(1).replace(' ', ''))
                lon = float(mlon.group(1).replace(' ', ''))
                if lon > 0:
                    lon = -lon
                return self.fmt.a_gms(lat), self.fmt.a_gms(lon)
            except ValueError:
                pass

        mlat = next((p.search(t) for p in P['lat_gms'] if p.search(t)), None)
        mlon = next((p.search(t) for p in P['lon_gms'] if p.search(t)), None)
        if mlat and mlon:
            lat = self.fmt.gms_str(mlat.group(1))
            lon = self.fmt.gms_str(mlon.group(1))
            if lon and not lon.startswith('-'):
                lon = f"-{lon}"
            return lat, lon

        return None, None

    def cambio(self, t: str) -> str:
        for m in P['cambio'].finditer(t):
            if m.group(0).upper().startswith('OC-'):
                v = m.group(0).upper()
            else:
                v = (m.group(1) or '').upper()
            if v and self._es_valido(v, 'CAMBIO'):
                return v
        return VACIO

    def etiqueta(self, t: str) -> str:
        lineas = t.splitlines()
        for linea in lineas:
            lu = linea.upper()
            if any(k in lu for k in LINEAS_ETIQUETA):
                m = P['etiqueta_todem_eol'].search(linea)
                if m:
                    v = m.group(0).strip()
                    if self._es_valido(v, 'ETIQUETA'):
                        return v
        for linea in lineas:
            if 'TO_DEM_' in linea.upper():
                m = P['etiqueta_todem_eol'].search(linea)
                if m:
                    v = m.group(0).strip()
                    if self._es_valido(v, 'ETIQUETA'):
                        return v
        m = P['etiqueta_todem'].search(t)
        if m:
            v = m.group(0)
            if self._es_valido(v, 'ETIQUETA'):
                return v
        return VACIO

    def extraer(self, t: str) -> Dict[str, Any]:
        lat, lon = self.coordenadas(t)
        coords = f"{lat} {lon}" if lat and lon else VACIO
        if coords != VACIO and not self._es_valido(coords, 'COORDENADAS'):
            coords = VACIO
        return {
            'ip_demarcador': self.ip_demarcador(t),
            'anillo': self.anillo(t),
            'ancho_banda': self.ancho_banda(t),
            'id_servicio': self.id_servicio(t),
            'latitud': lat,
            'longitud': lon,
            'coordenadas': coords,
            'tipo_servicio': self.tipo_servicio(t),
            'equipo_central': self.equipo_central(t),
            'puerto_common': self.puerto_common(t),
            'puerto_owner': self.puerto_owner(t),
            'cambio': self.cambio(t),
            'etiqueta': self.etiqueta(t),
        }


ORDEN_COLUMNAS = [
    'Equipo_Central', 'Puerto_Common', 'Puerto_Owner', 'Anillo', 'Etiqueta',
    'ID_Servicio', 'CAMBIO', 'IP_Demarcador', 'Archivo', 'Coordenadas',
]


class FileProcessor:
    def __init__(self, ruta_origen: Path):
        self.ruta_origen = ruta_origen
        self.extractor = DataExtractor()

    def obtener_archivos(self) -> List[Path]:
        if not self.ruta_origen.exists():
            return []
        return [
            f for f in self.ruta_origen.iterdir()
            if f.suffix.lower() == EXTENSION and f.is_file()
        ]

    def procesar(self) -> pd.DataFrame:
        archivos = self.obtener_archivos()
        if not archivos:
            return pd.DataFrame()

        filas = []
        for archivo in archivos:
            try:
                with open(archivo, 'r', encoding='utf-8', errors='ignore') as f:
                    texto = f.read()
            except Exception:
                continue

            d = self.extractor.extraer(texto)
            filas.append({
                'Archivo': archivo.stem,
                'ID_Servicio': d['id_servicio'],
                'Anillo': d['anillo'],
                'IP_Demarcador': d['ip_demarcador'],
                'ancho_banda': d['ancho_banda'],
                'Coordenadas': d['coordenadas'],
                'Tipo_Servicio': d['tipo_servicio'],
                'Equipo_Central': d['equipo_central'],
                'Puerto_Common': d['puerto_common'],
                'Puerto_Owner': d['puerto_owner'],
                'CAMBIO': d['cambio'],
                'Etiqueta': d['etiqueta'],
            })

        if not filas:
            return pd.DataFrame()

        df = pd.DataFrame(filas)
        cols = [c for c in ORDEN_COLUMNAS if c in df.columns]
        resto = [c for c in df.columns if c not in cols]
        return df[cols + resto]


class ExcelWriter:
    def __init__(self, ruta_destino: Path):
        self.ruta_destino = ruta_destino
        self.ruta_destino.mkdir(parents=True, exist_ok=True)

    def _ajustar_columnas(self, hoja):
        for col in hoja.columns:
            letra = col[0].column_letter
            max_len = max(
                (len(str(c.value)) for c in col if c.value is not None),
                default=0,
            )
            hoja.column_dimensions[letra].width = min(max_len + 2, 50)

    def _guardar_xlsx(self, df: pd.DataFrame, ruta: Path, engine: str) -> bool:
        try:
            with pd.ExcelWriter(ruta, engine=engine) as writer:
                df.to_excel(writer, index=False, sheet_name='DATOS')
                if engine == 'openpyxl':
                    self._ajustar_columnas(writer.sheets['DATOS'])
            return True
        except Exception:
            return False

    def _guardar_xls(self, df: pd.DataFrame, ruta: Path) -> bool:
        try:
            with pd.ExcelWriter(ruta, engine='xlwt') as writer:
                df.to_excel(writer, index=False, sheet_name='DATOS')
            return True
        except Exception:
            return False

    def guardar(self, df: pd.DataFrame) -> Optional[Path]:
        if df.empty:
            return None

        ruta_xlsx = self.ruta_destino / NOMBRE_SALIDA

        if self._guardar_xlsx(df, ruta_xlsx, 'openpyxl'):
            return ruta_xlsx

        if self._guardar_xlsx(df, ruta_xlsx, 'xlsxwriter'):
            return ruta_xlsx

        ruta_xls = ruta_xlsx.with_suffix('.xls')
        if self._guardar_xls(df, ruta_xls):
            return ruta_xls

        ruta_csv = ruta_xlsx.with_suffix('.csv')
        df.to_csv(ruta_csv, index=False, encoding='utf-8-sig')
        return ruta_csv


def main():
    if not RUTA_ORIGEN.exists():
        print(f"[ERROR] Ruta de origen no existe: {RUTA_ORIGEN}")
        return

    print("Procesando archivos...")
    processor = FileProcessor(RUTA_ORIGEN)
    df = processor.procesar()

    if df.empty:
        print("[AVISO] No se encontraron datos.")
        return

    print(f"[OK] {len(df)} archivos procesados.")
    writer = ExcelWriter(RUTA_DESTINO)
    salida = writer.guardar(df)

    if salida and salida.suffix.lower() in ('.xlsx', '.xls'):
        print(f"[OK] Excel guardado en: {salida}")
    elif salida:
        print(f"[AVISO] No se pudo generar Excel. Se guardo como: {salida}")
    else:
        print("[ERROR] No se pudo guardar el archivo.")


if __name__ == "__main__":
    main()