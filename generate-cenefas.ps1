# Script legado del prototipo anterior.
# El flujo actual recomendado es la app web:
# 1. npm.cmd install
# 2. npm.cmd start
# 3. abrir http://localhost:3000

param(
  [string]$ExcelPath = "Listado dia de la pizza 1.xlsx",
  [string]$OutputPath = "index.html"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-SharedStrings {
  param([System.IO.Compression.ZipArchive]$Zip)

  $entry = $Zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
  if (-not $entry) {
    return @()
  }

  $reader = [System.IO.StreamReader]::new($entry.Open())
  try {
    $xml = [xml]$reader.ReadToEnd()
  }
  finally {
    $reader.Close()
  }

  $values = New-Object System.Collections.Generic.List[string]
  foreach ($si in $xml.sst.si) {
    $values.Add($si.InnerText)
  }
  return $values
}

function Get-CellValue {
  param(
    $Cell,
    [string[]]$SharedStrings
  )

  if (-not $Cell) {
    return ""
  }

  $cellType = if ($Cell.PSObject.Properties["t"]) { [string]$Cell.t } else { "" }

  if ($cellType -eq "s") {
    return $SharedStrings[[int]$Cell.v]
  }

  if ($cellType -eq "inlineStr") {
    return [string]$Cell.InnerText
  }

  if ($Cell.v) {
    return [string]$Cell.v
  }

  return ""
}

function Format-Price {
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return ""
  }

  $number = [double]::Parse($Value, [System.Globalization.CultureInfo]::InvariantCulture)
  return $number.ToString("0.##", [System.Globalization.CultureInfo]::GetCultureInfo("es-UY"))
}

function Format-Discount {
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return ""
  }

  if ($Value.Contains("%")) {
    return $Value
  }

  $number = [double]::Parse($Value, [System.Globalization.CultureInfo]::InvariantCulture)
  if ($number -le 1) {
    $number *= 100
  }

  return ($number.ToString("0.##", [System.Globalization.CultureInfo]::GetCultureInfo("es-UY")) + "%")
}

function Encode-Html {
  param([string]$Value)
  return [System.Net.WebUtility]::HtmlEncode($Value)
}

function Get-Products {
  param([string]$Path)

  $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
  try {
    $sharedStrings = Get-SharedStrings -Zip $zip
    $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet1.xml" }
    if (-not $sheetEntry) {
      throw "No se encontro la hoja xl/worksheets/sheet1.xml en el archivo Excel."
    }

    $reader = [System.IO.StreamReader]::new($sheetEntry.Open())
    try {
      $sheetXml = [xml]$reader.ReadToEnd()
    }
    finally {
      $reader.Close()
    }

    $namespace = [System.Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
    $namespace.AddNamespace("a", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    $rows = $sheetXml.SelectNodes("//a:sheetData/a:row", $namespace)
    $products = New-Object System.Collections.Generic.List[object]

    foreach ($row in ($rows | Select-Object -Skip 1)) {
      $map = @{}
      foreach ($cell in $row.c) {
        $column = ([string]$cell.r) -replace "\d", ""
        $map[$column] = Get-CellValue -Cell $cell -SharedStrings $sharedStrings
      }

      if ([string]::IsNullOrWhiteSpace($map["A"]) -and [string]::IsNullOrWhiteSpace($map["B"])) {
        continue
      }

      $products.Add([pscustomobject]@{
        Codigo       = [string]$map["A"]
        Descripcion  = [string]$map["B"]
        PrecioVta    = Format-Price -Value ([string]$map["C"])
        Accion       = Format-Discount -Value ([string]$map["D"])
        PrecioOferta = Format-Price -Value ([string]$map["E"])
      })
    }

    return $products
  }
  finally {
    $zip.Dispose()
  }
}

function New-BlockMarkup {
  param($Product)

  $codigo = Encode-Html $Product.Codigo
  $descripcion = Encode-Html $Product.Descripcion
  $precioVta = Encode-Html $Product.PrecioVta
  $precioOferta = Encode-Html $Product.PrecioOferta
  $accion = Encode-Html $Product.Accion

  return @"
      <section class="block">
        <div class="block__left">
          <div class="image-slot">
            <span class="image-slot__label">Imagen del producto</span>
            <span class="image-slot__code">COD $codigo</span>
          </div>
        </div>
        <div class="block__right">
          <article class="product-card">
            <p class="product-card__code">COD : $codigo</p>
            <h2 class="product-card__description">$descripcion</h2>
            <div class="product-card__divider"></div>
            <p class="product-card__regular">PRECIO REGULAR: U`$S $precioVta</p>
            <p class="product-card__offer-label">OFERTA</p>
            <p class="product-card__offer">U`$S $precioOferta</p>
            <p class="product-card__discount">Descuento: $accion</p>
          </article>
        </div>
      </section>
"@
}

function New-PageMarkup {
  param([object[]]$Products)

  $blocks = ($Products | ForEach-Object { New-BlockMarkup -Product $_ }) -join "`r`n"
  return @"
  <section class="page">
$blocks
  </section>
"@
}

$excelFullPath = Join-Path (Get-Location) $ExcelPath
$outputFullPath = Join-Path (Get-Location) $OutputPath
$products = Get-Products -Path $excelFullPath

$pages = New-Object System.Collections.Generic.List[string]
for ($index = 0; $index -lt $products.Count; $index += 3) {
  $lastIndex = [Math]::Min($index + 2, $products.Count - 1)
  $group = for ($cursor = $index; $cursor -le $lastIndex; $cursor++) { $products[$cursor] }
  $pages.Add((New-PageMarkup -Products $group))
}

$pagesMarkup = $pages -join "`r`n"
$html = @"
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Cenefas Dia de la Pizza</title>
  <style>
    :root {
      --page-width: 210mm;
      --page-height: 297mm;
      --gap-size: 1cm;
      --sheet-bg: #bfefff;
      --border-color: #51423f;
      --left-fill: #b7b7b7;
      --right-fill: #f8d7e6;
      --block-height: calc((var(--page-height) - (2 * var(--gap-size))) / 3);
      --text-dark: #4a3a35;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      background: #e8f4fb;
      color: var(--text-dark);
      font-family: "Arial Narrow", Arial, sans-serif;
    }

    .document {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 10mm;
      padding: 10mm 0;
    }

    .page {
      width: var(--page-width);
      min-height: var(--page-height);
      display: flex;
      flex-direction: column;
      gap: var(--gap-size);
      padding: 0;
      background: var(--sheet-bg);
      page-break-after: always;
      overflow: hidden;
    }

    .page:last-child {
      page-break-after: auto;
    }

    .block {
      height: var(--block-height);
      display: grid;
      grid-template-columns: 1fr 2fr;
      border: 0.8mm solid var(--border-color);
    }

    .block__left {
      background: var(--left-fill);
      border-right: 0.8mm solid var(--border-color);
      padding: 6mm;
    }

    .image-slot {
      width: 100%;
      height: 100%;
      border: 0.6mm dashed rgba(81, 66, 63, 0.7);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 4mm;
      text-align: center;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: rgba(81, 66, 63, 0.82);
      background:
        linear-gradient(135deg, rgba(255, 255, 255, 0.3), transparent),
        rgba(255, 255, 255, 0.12);
    }

    .image-slot__label {
      font-size: 5mm;
      font-weight: 700;
    }

    .image-slot__code {
      font-size: 3.5mm;
      font-weight: 600;
    }

    .block__right {
      background: var(--right-fill);
      padding: 5mm 6mm 4mm;
    }

    .product-card {
      height: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
      text-align: center;
      justify-content: flex-start;
    }

    .product-card__code {
      margin: 0;
      font-size: 3.2mm;
      font-weight: 700;
      letter-spacing: 0.06em;
    }

    .product-card__description {
      margin: 2mm 0 3mm;
      font-size: 8mm;
      line-height: 0.98;
      font-weight: 900;
      text-transform: uppercase;
      max-width: 95%;
    }

    .product-card__divider {
      width: 100%;
      border-top: 0.6mm solid rgba(81, 66, 63, 0.75);
      margin: 1mm 0 3mm;
    }

    .product-card__regular {
      margin: 0;
      font-size: 5mm;
      font-weight: 700;
      text-transform: uppercase;
    }

    .product-card__offer-label {
      margin: 2mm 0 0;
      font-size: 9mm;
      font-weight: 900;
      line-height: 1;
      text-transform: uppercase;
    }

    .product-card__offer {
      margin: 0;
      font-size: 16mm;
      line-height: 1;
      font-weight: 900;
      text-transform: uppercase;
    }

    .product-card__discount {
      margin: auto 0 0;
      font-size: 3.5mm;
      font-weight: 700;
      text-transform: uppercase;
    }

    @page {
      size: A4;
      margin: 0;
    }

    @media print {
      body {
        background: transparent;
      }

      .document {
        gap: 0;
        padding: 0;
      }

      .page {
        margin: 0;
      }
    }
  </style>
</head>
<body>
  <main class="document">
$pagesMarkup
  </main>
</body>
</html>
"@

[System.IO.File]::WriteAllText($outputFullPath, $html, [System.Text.Encoding]::UTF8)
Write-Output "Generado $OutputPath con $($products.Count) productos en $($pages.Count) hojas."
