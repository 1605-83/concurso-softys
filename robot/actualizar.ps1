# Robot de actualizacion de los tableros (Ayudin, Softys, Georgalos).
# Lee la base de Gescom (Lago Puelo + Elebes) y arma el mismo CSV que el
# reporte "Detallado de ventas extendido" que antes se bajaba a mano.
#
# Uso:  powershell -File robot\actualizar.ps1 [-Repo Ventas-Softys] [-Salida <carpeta>]
# La clave se lee de la variable de entorno GESCOM_LPE_CLAVE (secreto de GitHub).
param(
  [string]$Repo = $(if ($env:GITHUB_REPOSITORY) { $env:GITHUB_REPOSITORY.Split('/')[-1] } else { '' }),
  [string]$Salida = '.'
)
$ErrorActionPreference = 'Stop'

# Que proveedores lleva cada tablero y cuantos meses (1 = mes en curso; 2 = el anterior tambien)
$CONFIG = @{
  'Ventas-Grupo-Ayudin' = @{ proveedores = @('247');      meses = 1 }   # GRUPO AYUDIN ARGENTINA S.A
  'Ventas-Softys'       = @{ proveedores = @('150');      meses = 1 }   # SOFTYS ARGENTINA S.A
  'Cobertura-Georgalos' = @{ proveedores = @('42','43');  meses = 2 }   # GEORGALOS + GENERAL CEREALS
  'Concurso-Bic'        = @{ proveedores = @('119');      meses = 1 }   # BIC ARGENTINA S.A
  'concurso-softys'     = @{ proveedores = @('150');      desde = '2026-09-01' }   # SOFTYS, desde el inicio del concurso
  # Reporte diario LP-ELB: los proveedores del nomenclador del Excel (por nombre), entregas del mes
  # incluidas las ya cargadas para los proximos dias, y un CSV liviano con las columnas que usa
  'Reporte-Diario-Ventas' = @{
    nombres = @('SOFTYS ARGENTINA S.A','GRUPO AYUDIN ARGENTINA S.A','ILOLAY','GEORGALOS','GENERAL CEREALS S.A','RIOSMA','CEPAS ARGENTINAS S.A',
                'AJINOMOTO','BETTER FOOD SAS','BIC ARGENTINA S.A','BODEGAS SAN HUBERTO S.A','BRURIN','CASTELL S.A','DREAMCO S.A','FECOVITA',
                'LABORATORIOS ECOVITA S.A','LCB','LEDESMA','MENOYO S.A.','MOLINO CHACABUCO S.A','MORIXE','NECHO S.A','POLDITOS S.A.S',
                'PORTA HNOS S.A','PRIMEROS PRODUCTOS PEHUENIA','PRO DE MAN S.A','INDUSTRIAS QUIMICAS Y MINERAS TIMBO S.A','LINEA DORADA S.A')
    desde = '2026-08-01'; porMes = $true; futuro = 7   # la base tiene ventas desde el 1/8/2026
    columnas = @('Cliente','FechaComprobante','FechaEntrega','NroComprobante','TipoDeVenta','Empresa','Codigo','CantBase','ImporteNetoItem',
                 'ImporteItem','RazonSocial','CodVendedor','Vendedor','Articulo','PrecioCosto','Proveedor','Categoria','FechaCarga','MotivoDevolucion')
  }
}
$cfg = $CONFIG[$Repo]
if (-not $cfg) { throw "Repo desconocido: '$Repo'. Opciones: $($CONFIG.Keys -join ', ')" }
$clave = $env:GESCOM_LPE_CLAVE
if (-not $clave) { throw 'Falta la variable de entorno GESCOM_LPE_CLAVE' }

$BASE = 'https://datos-gescom.panelempresas.workers.dev'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Pedir([string]$metodo, [string]$ruta, $cuerpo) {
  $p = @{ Uri = "$BASE$ruta"; Method = $metodo; UseBasicParsing = $true; Headers = @{ Authorization = "Bearer $clave" } }
  if ($cuerpo) { $p.ContentType = 'application/json'; $p.Body = [Text.Encoding]::UTF8.GetBytes(($cuerpo | ConvertTo-Json -Compress)) }
  $r = Invoke-WebRequest @p
  $txt = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
  $j = $txt | ConvertFrom-Json
  if ($j.error) { throw "La base respondio error: $($j.error)" }
  $j
}

# ---- Periodo: por FECHA DE COMPROBANTE, del 1 del mes a hoy (hora argentina) ----
$hoy = (Get-Date).ToUniversalTime().AddHours(-3).Date
if ($cfg.desde) { $desde = [datetime]::ParseExact($cfg.desde, 'yyyy-MM-dd', $null) }
else { $desde = (Get-Date -Year $hoy.Year -Month $hoy.Month -Day 1).Date.AddMonths(1 - $cfg.meses) }

# "Datos al": la ultima actualizacion de ventas de la base (UTC -> Argentina)
$corte = (Get-Date).ToUniversalTime().AddHours(-3)
try {
  $est = Pedir 'GET' '/estado'
  $v = $est.datos | Where-Object { $_.que -eq 'ventas' } | Select-Object -First 1
  if ($v.actualizado) { $corte = ([datetime]::Parse($v.actualizado, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal)).AddHours(-3) }
} catch { Write-Warning "No se pudo leer /estado: $_" }

# Todos los tableros toman la venta por FECHA DE COMPROBANTE (pedido de Bruno, 9/10/2026): solo lo facturado, hasta hoy
$hasta = $hoy

function Consultar([datetime]$d1, [datetime]$d2) {
  if ($cfg.nombres) {
    $lista = ($cfg.nombres | ForEach-Object { "'" + $_.Replace("'", "''") + "'" }) -join ','
    $prov = "SELECT codigo FROM proveedores WHERE nombre IN ($lista)"
  } else { $prov = ($cfg.proveedores | ForEach-Object { "'$_'" }) -join ',' }
  $sql = @"
SELECT v.id, v.fecha, COALESCE(NULLIF(v.entrega,''), v.fecha) AS entrega, v.tipo,
       -- FECHA DE COMPROBANTE: la base no la guarda. Medido contra el reporte de Gescom (12.411 renglones, oct 2026):
       -- facturas y canjes = fecha de entrega (99,4%); NC por rechazo y notas de debito = dia de carga de la nota (100%)
       (CASE WHEN v.tipo IN ('DEV-RE','DEB') THEN v.fecha ELSE COALESCE(NULLIF(v.entrega,''), v.fecha) END) AS fecha_comp, v.empresa, v.cliente,
       v.vendedor, v.reparto, v.chofer, COALESCE(ch.nombre, v.chofer) AS chofer_nombre, v.comprobante,
       v.ref_id, v.directa, v.origen, v.motivo,
       i.articulo, i.cantidad, i.factor, i.neto, i.total, i.precio_costo, i.precio_unitario,
       a.descripcion AS articulo_nombre, p.nombre AS proveedor_nombre,
       c.razon_social, c.nombre AS cliente_nombre, c.localidad, c.ramo, sr.descripcion AS subramo,
       c.condicion_pago, c.lista_precio,
       ve.nombre AS vendedor_nombre, e.superior AS cod_supervisor, su.nombre AS supervisor
FROM ventas v
JOIN venta_items i ON i.venta_id = v.id
JOIN articulos a ON a.codigo = i.articulo
LEFT JOIN proveedores p ON p.codigo = a.proveedor
LEFT JOIN clientes c ON c.codigo = v.cliente
LEFT JOIN subramos sr ON sr.codigo = c.subramo
LEFT JOIN vendedores ve ON ve.codigo = v.vendedor
LEFT JOIN empleados e ON e.codigo = v.vendedor
LEFT JOIN empleados su ON su.codigo = e.superior
LEFT JOIN empleados ch ON ch.codigo = v.chofer
WHERE a.proveedor IN ($prov)
  AND v.tipo IN ('VEN','DEB','DEV-RE','DEV-CA')
  AND NULLIF(v.comprobante,'') IS NOT NULL
  AND (CASE WHEN v.tipo IN ('DEV-RE','DEB') THEN v.fecha ELSE COALESCE(NULLIF(v.entrega,''), v.fecha) END)
      BETWEEN '$($d1.ToString('yyyy-MM-dd'))' AND '$($d2.ToString('yyyy-MM-dd'))'
ORDER BY v.id, i.orden
"@
  $r = Pedir 'POST' '/consulta' @{ sql = $sql }
  if ($r.truncado) {
    # La base devuelve hasta 20.000 filas: si se corta, se pide en dos mitades
    if ($d1 -ge $d2) { throw "La consulta del $($d1.ToString('yyyy-MM-dd')) se corto en $($r.cantidad) filas (tope de la base)" }
    $medio = $d1.AddDays([math]::Floor(($d2 - $d1).TotalDays / 2))
    return @(Consultar $d1 $medio) + @(Consultar $medio.AddDays(1) $d2)
  }
  @($r.filas)
}

if (-not $cfg.porMes) {
  $filas = Consultar $desde $hasta
  # El dia 1 a primera hora el mes nuevo puede estar vacio: en ese caso se muestra el mes anterior entero
  if ($filas.Count -eq 0) {
    $desde = $desde.AddMonths(-1); $hasta = $desde.AddMonths($cfg.meses).AddDays(-1)
    Write-Warning "Sin ventas desde el 1 del mes; uso $($desde.ToString('yyyy-MM-dd')) a $($hasta.ToString('yyyy-MM-dd'))"
    $filas = Consultar $desde $hasta
  }
  if ($filas.Count -eq 0) { throw 'La base no devolvio ninguna venta para el periodo: no se reemplaza el CSV' }
}

# ---- Datos que la base no tiene (Familia del articulo y direccion del cliente) ----
$extra = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'datos.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json

# ---- Armado del CSV, mismas 62 columnas que el reporte de Gescom ----
$COLS = 'Cliente;Direccion;Localidad;FechaComprobante;FechaEntrega;FechaCarga;NroComprobante;TipoDeVenta;Empresa;Codigo;CantBase;ImporteNetoItem;ImporteItem;RazonSocial;MotivoDevolucion;Descuento;CodVendedor;Vendedor;RutaPreventa;Articulo;NetoItem;Reparto;ComentarioInterno;Ramo;Subramo;PrecioCosto;Chofer;valorDescuento;Mon_Simbolo;Mon_Decimales;Proveedor;PesoKg;ImpuestoInterno;CondicionPago;FechaLiquidacion;ComboCodigo;Marca;Linea;Sabor;Calibre;Familia;Rubro;Tags;Promociones;SegmentoRentabilidad;CodSupervisor;NomSupervisor;Origen;ComprobanteReferencia;Etiqueta;VentaDirecta;NumeroVenta;Pendiente;ListaPrecio;Taxonomia;OrdenPreparacion;FechaPreparacion;Bodeguistas;EtiquetaItem;UnidadFactor;PesoKgReal;NetoItemReal'.Split(';')
if ($cfg.columnas) { $COLS = $cfg.columnas }   # CSV liviano: solo las columnas que usa ese tablero
$EMPRESAS = @{ '3' = 'LAGOPUELO S.A'; '97' = 'Empresa LAGOPUELO'; '1' = 'ELEBES S.A.'; '99' = 'ELEBES' }
$o = [char]0xF3; $e = [char]0xE9   # o y e con acento (el script queda en ASCII puro)
$TIPOS = @{ 'VEN' = 'Venta'; 'DEB' = "Nota de D$($e)bito"; 'DEV-RE' = "Devoluci$($o)n por Rechazo"; 'DEV-CA' = "Devoluci$($o)n por Canje" }
$ar = [Globalization.CultureInfo]::GetCultureInfo('es-AR')
$TI = $ar.TextInfo
$SIN_CAT = "Sin categor$([char]0xED)a"

function Num($x, [int]$dec = 2) { if ($null -eq $x -or "$x" -eq '') { return '' }; ([double]$x).ToString("0.$('0' * $dec)", $ar).Replace('.', '') }
function Dmy([string]$iso) { if ($iso -match '^(\d{4})-(\d{2})-(\d{2})') { "$($Matches[3])-$($Matches[2])-$($Matches[1])" } else { '' } }
function Txt($x) {
  $t = "$x".Trim() -replace '[\r\n]+', ' '
  if ($t -match '[;"]') { '"' + $t.Replace('"', '""') + '"' } else { $t }
}

function ArmarCsv($filas) {
$sb = New-Object Text.StringBuilder
[void]$sb.Append(($COLS -join ';') + "`r`n")
foreach ($f in $filas) {
  $signo = if ($f.tipo -like 'DEV-*') { -1 } else { 1 }
  $unid = [double]$f.cantidad * $(if ($f.factor) { [double]$f.factor } else { 1 })
  $pend = [string]::IsNullOrWhiteSpace($f.comprobante)
  $desc = ''
  if ($f.precio_unitario -and $f.cantidad -and [double]$f.precio_unitario * [double]$f.cantidad -ne 0) {
    $desc = (Num (100 * (1 - [double]$f.neto / ([double]$f.precio_unitario * [double]$f.cantidad)))) + ' %'
  }
  $art = "$($f.articulo)"; $cli = "$($f.cliente)".Trim()
  $r = [ordered]@{}
  $r.Cliente = $cli
  $r.Direccion = $extra.direcciones.$cli
  $r.Localidad = $f.localidad
  $r.FechaComprobante = if ($pend) { '' } else { $f.fecha_comp }
  $r.FechaEntrega = Dmy $f.entrega
  $r.FechaCarga = Dmy $f.fecha
  $r.NroComprobante = if ($pend) { "VEN-$($f.id)" } else { "$($f.comprobante)-$($f.id)" }
  $r.TipoDeVenta = $TIPOS["$($f.tipo)"]
  $r.Empresa = if ($EMPRESAS["$($f.empresa)"]) { $EMPRESAS["$($f.empresa)"] } else { "$($f.empresa)" }
  $r.Codigo = $art
  $r.CantBase = Num ($signo * $unid)
  $r.ImporteNetoItem = Num ($signo * [double]$f.neto)
  $r.ImporteItem = Num ($signo * [double]$f.total)
  $r.RazonSocial = if ($f.razon_social) { $f.razon_social } else { $f.cliente_nombre }
  $r.MotivoDevolucion = $f.motivo
  $r.Descuento = $desc
  $r.CodVendedor = $f.vendedor
  $r.Vendedor = $f.vendedor_nombre
  $r.Articulo = $f.articulo_nombre
  $r.NetoItem = if ($unid -ne 0) { Num ([double]$f.neto / $unid) } else { '' }
  $r.Reparto = $f.reparto
  $r.Ramo = $f.ramo
  $r.Subramo = $f.subramo
  $r.PrecioCosto = Num ($signo * [double]$f.precio_costo * [double]$f.cantidad)   # costo del renglon, como el reporte de Gescom
  $r.Chofer = $f.chofer_nombre
  $r.Mon_Simbolo = '$'
  $r.Mon_Decimales = '2'
  $r.Proveedor = $f.proveedor_nombre
  $r.CondicionPago = $f.condicion_pago
  $r.Familia = $extra.familias.$art
  # Categoria de producto (sacada del reporte de Gescom; la base no la tiene)
  $cat = $extra.categorias.$art; $provNom = "$($f.proveedor_nombre)".Trim()
  if (-not $cat -or $extra.catGenericas -contains $cat) {
    # articulo nuevo o con categoria generica: reglas por nombre del articulo
    $artNom = "$($f.articulo_nombre)".ToUpper()
    foreach ($rg in $extra.catReglas) { if ($rg.p -eq $provNom -and $artNom -match $rg.r) { $cat = $rg.c; break } }
  }
  if (-not $cat -and $r.Familia) { $cat = $TI.ToTitleCase("$($r.Familia)".ToLower()) }
  if (-not $cat) { $cat = $extra.catProveedor.$provNom }   # proveedor con una sola categoria
  # Categorias finales de cada empresa (lista cerrada): reglas que mandan, renombres, y lo que no es valido va por reglas o al defecto
  $fin = $extra.catFinal.$provNom
  if ($fin) {
    $artNom = "$($f.articulo_nombre)".ToUpper()
    $forzada = $null; foreach ($rg in $fin.forzar) { if ($artNom -match $rg[0]) { $forzada = $rg[1]; break } }
    if ($forzada) { $cat = $forzada }
    else {
      if ($cat -and $fin.renombrar.PSObject.Properties.Name -contains $cat) { $cat = $fin.renombrar.$cat }
      if ($fin.validas -notcontains $cat) {
        $cat = $null; foreach ($rg in $fin.reglas) { if ($artNom -match $rg[0]) { $cat = $rg[1]; break } }
        if (-not $cat) { $cat = $fin.defecto }
      }
    }
  }
  if (-not $cat -and $provNom) { $cat = $TI.ToTitleCase($provNom.ToLower()) }   # ultimo recurso: el proveedor
  $r.Categoria = if ($cat) { $cat } else { $SIN_CAT }
  $r.CodSupervisor = $f.cod_supervisor
  $r.NomSupervisor = $f.supervisor
  $r.Origen = $f.origen
  $r.ComprobanteReferencia = $f.ref_id
  $r.VentaDirecta = if ("$($f.directa)" -eq '1') { 'True' } else { 'False' }
  $r.NumeroVenta = $f.id
  $r.Pendiente = if ($pend) { '1' } else { '0' }
  $r.ListaPrecio = $f.lista_precio
  $r.UnidadFactor = if ($f.factor) { Num $f.factor 0 } else { '1' }
  $r.NetoItemReal = $r.NetoItem
  [void]$sb.Append((($COLS | ForEach-Object { Txt $r[$_] }) -join ';') + "`r`n")
}
$sb.ToString()
}

if ($cfg.porMes) {
  # ---- Historial: un CSV por mes en la carpeta meses/. Se rehacen el mes en curso y el anterior
  # (devoluciones y correcciones tardias); los meses mas viejos ya guardados no se vuelven a bajar ----
  $dir = Join-Path (Resolve-Path $Salida) 'meses'
  New-Item -ItemType Directory -Force $dir | Out-Null
  # hora de los datos + hora de la corrida: cada corrida deja un nombre nuevo y nadie recibe una copia vieja del CSV
  $sello = $corte.ToString('yyyyMMdd-HHmmss') + '-r' + (Get-Date).ToUniversalTime().ToString('yyyyMMddHHmmss')
  $mesActual = (Get-Date -Year $hoy.Year -Month $hoy.Month -Day 1).Date
  $m = (Get-Date -Year $desde.Year -Month $desde.Month -Day 1).Date
  while ($m -le $mesActual) {
    $mesClave = $m.ToString('yyyy-MM')
    $previos = @(Get-ChildItem -Path $dir -Filter "ventas-$mesClave-*.csv" -File)
    if ($m -lt $mesActual.AddMonths(-1) -and $previos.Count) { Write-Host "$mesClave : ya guardado"; $m = $m.AddMonths(1); continue }
    # corridas de cada hora: el mes anterior solo se rehace en la primera del dia (antes de las 9)
    $horaArg = (Get-Date).ToUniversalTime().AddHours(-3).Hour
    if ($m -lt $mesActual -and $previos.Count -and $horaArg -ge 9) { Write-Host "$mesClave : se rehace en la corrida de la manana"; $m = $m.AddMonths(1); continue }
    $fin = $m.AddMonths(1).AddDays(-1); if ($fin -gt $hasta) { $fin = $hasta }
    $filasMes = @(Consultar $m $fin)
    if ($filasMes.Count -eq 0) { Write-Host "$mesClave : sin ventas todavia"; $m = $m.AddMonths(1); continue }
    $destino = Join-Path $dir "ventas-$mesClave-$sello.csv"
    [IO.File]::WriteAllText($destino, (ArmarCsv $filasMes), [Text.Encoding]::GetEncoding(1252))
    $previos | Where-Object { $_.FullName -ne $destino } | Remove-Item -Force
    Write-Host "$Repo $mesClave : $($filasMes.Count) renglones, datos al $($corte.ToString('dd/MM/yyyy HH:mm')) -> $(Split-Path $destino -Leaf)"
    $m = $m.AddMonths(1)
  }
  # los CSV sueltos de la raiz ya no se usan en este tablero
  Get-ChildItem -Path $Salida -Filter '*.csv' -File | Remove-Item -Force
  return
}

# ---- Guardar: reemplaza el CSV anterior (los tableros leen el mas nuevo de la raiz del repo) ----
$nombre = "ventas-Detallado de ventas extendido-$($corte.ToString('yyyyMMdd-HHmmss')).csv"
$destino = Join-Path (Resolve-Path $Salida) $nombre
Get-ChildItem -Path $Salida -Filter '*.csv' -File | Where-Object { $_.FullName -ne $destino } | Remove-Item -Force
[IO.File]::WriteAllText($destino, (ArmarCsv $filas), [Text.Encoding]::GetEncoding(1252))

$n = ($filas | Measure-Object).Count
$neto = ($filas | ForEach-Object { if ($_.tipo -like 'DEV-*') { -[double]$_.neto } else { [double]$_.neto } } | Measure-Object -Sum).Sum
Write-Host "$Repo : $n renglones, comprobantes del $($desde.ToString('dd/MM/yyyy')) a $($hasta.ToString('dd/MM/yyyy')),venta neta $([math]::Round($neto).ToString('N0', $ar)) (sin IVA), datos al $($corte.ToString('dd/MM/yyyy HH:mm'))"
Write-Host "Archivo: $nombre"
