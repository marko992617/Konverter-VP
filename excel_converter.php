<?php
// Excel konverter sa TIP1, TIP2 i TIP3 u import template sa generisanjem "Broj pošiljke"
// Dodati partneri CS, MB, PM, SL, EP, SV i izlaz u .xls
require __DIR__ . '/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

function findColumnIndex(Worksheet $sheet, string $headerName): ?int {
    $lastCol = Coordinate::columnIndexFromString($sheet->getHighestColumn());
    for ($i = 1; $i <= $lastCol; $i++) {
        $v = trim((string)$sheet->getCellByColumnAndRow($i, 1)->getValue());
        if (mb_strtolower($v) === mb_strtolower($headerName)) {
            return $i;
        }
    }
    return null;
}

// Partneri i sekvence
$partners = ['CS','MB','PM','SL','EP','SV'];
$seqFile = __DIR__ . '/sequences.json';
$seqData = file_exists($seqFile)
    ? json_decode(file_get_contents($seqFile), true)
    : array_fill_keys($partners, 0);

$resultHtml = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $partner = $_POST['partner'] ?? '';
    if (!in_array($partner, $partners, true)) {
        $resultHtml = '<p class="error">Izaberite validnog partnera.</p>';
    } else {
        $prefix = $partner . '2';
        $tmpFile = $_FILES['excel_file']['tmp_name'] ?? null;
        $origName = pathinfo($_FILES['excel_file']['name'] ?? '', PATHINFO_FILENAME);
        if (!$tmpFile) {
            $resultHtml = '<p class="error">Fajl nije poslat.</p>';
        } else {
            $input = IOFactory::load($tmpFile);
            $inSheet = $input->getActiveSheet();
            $rows = $inSheet->getHighestRow();
            $tplPath = __DIR__ . '/Template Import PHP.xls';
            if (!file_exists($tplPath)) {
                $resultHtml = '<p class="error">Template nije pronađen.</p>';
            } else {
                $tpl = IOFactory::load($tplPath);
                $ts = $tpl->getActiveSheet();
                $tplCols = ['Primalac','Primalac adresa naziv','Primalac telefon','OTKUP','NAPOMENA','Broj pošiljke'];
                foreach ($tplCols as $col) {
                    $idxTpl[$col] = findColumnIndex($ts, $col) ?: exit("<p class='error'>Template nema kolonu {$col}.</p>");
                }
                // Indeksi ulaznih kolona
                $inHeaders = [
                    'Naziv','Adresa','Telefon','Otkup',
                    'Ime','Prezime','Adresa','Grad','Za naplatu','Proizvodi',
                    // TIP3
                    'Naziv primaoca ','Adresa i deo grada','Napomena'
                ];
                foreach ($inHeaders as $h) {
                    $inIdx[$h] = findColumnIndex($inSheet, $h);
                }
                // Sekvenca
                $seq = $seqData[$partner];
                $outRow = 2;
                for ($r = 2; $r <= $rows; $r++, $outRow++) {
                    $mapped = false;
                    // TIP1
                    if ($inIdx['Naziv'] && $inIdx['Adresa'] && $inIdx['Telefon'] && $inIdx['Otkup']) {
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac'],$outRow,$inSheet->getCellByColumnAndRow($inIdx['Naziv'],$r)->getValue());
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac adresa naziv'],$outRow,$inSheet->getCellByColumnAndRow($inIdx['Adresa'],$r)->getValue());
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac telefon'],$outRow,$inSheet->getCellByColumnAndRow($inIdx['Telefon'],$r)->getValue());
                        $ts->setCellValueByColumnAndRow($idxTpl['OTKUP'],$outRow,$inSheet->getCellByColumnAndRow($inIdx['Otkup'],$r)->getValue());
                        $ts->setCellValueByColumnAndRow($idxTpl['NAPOMENA'],$outRow,'');
                        $mapped = true;
                    }
                    // TIP2
                    elseif ($inIdx['Ime'] && $inIdx['Prezime'] && $inIdx['Za naplatu']) {
                        $fn = trim($inSheet->getCellByColumnAndRow($inIdx['Ime'],$r)->getValue().' '.$inSheet->getCellByColumnAndRow($inIdx['Prezime'],$r)->getValue());
                        $ad = trim($inSheet->getCellByColumnAndRow($inIdx['Adresa'],$r)->getValue().' '.$inSheet->getCellByColumnAndRow($inIdx['Grad'],$r)->getValue());
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac'],$outRow,$fn);
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac adresa naziv'],$outRow,$ad);
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac telefon'],$outRow,$inIdx['Telefon']?$inSheet->getCellByColumnAndRow($inIdx['Telefon'],$r)->getValue():'');
                        $ts->setCellValueByColumnAndRow($idxTpl['OTKUP'],$outRow,$inSheet->getCellByColumnAndRow($inIdx['Za naplatu'],$r)->getValue());
                        $ts->setCellValueByColumnAndRow($idxTpl['NAPOMENA'],$outRow,$inIdx['Proizvodi']?$inSheet->getCellByColumnAndRow($inIdx['Proizvodi'],$r)->getValue():'');
                        $mapped = true;
                    }
                    // TIP3
                    elseif ($inIdx['Naziv primaoca '] && $inIdx['Adresa i deo grada']) {
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac'],$outRow,$inSheet->getCellByColumnAndRow($inIdx['Naziv primaoca '],$r)->getValue());
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac adresa naziv'],$outRow,$inSheet->getCellByColumnAndRow($inIdx['Adresa i deo grada'],$r)->getValue());
                        $ts->setCellValueByColumnAndRow($idxTpl['Primalac telefon'],$outRow,$inIdx['Telefon']?$inSheet->getCellByColumnAndRow($inIdx['Telefon'],$r)->getValue():'');
                        $ts->setCellValueByColumnAndRow($idxTpl['OTKUP'],$outRow,$inIdx['Otkup']?$inSheet->getCellByColumnAndRow($inIdx['Otkup'],$r)->getValue():'');
                        $ts->setCellValueByColumnAndRow($idxTpl['NAPOMENA'],$outRow,$inIdx['Napomena']?$inSheet->getCellByColumnAndRow($inIdx['Napomena'],$r)->getValue():'');
                        $mapped = true;
                    }
                    if ($mapped) {
                        $seq++;
                        $num = $prefix . str_pad($seq,8,'0',STR_PAD_LEFT);
                        $ts->setCellValueByColumnAndRow($idxTpl['Broj pošiljke'],$outRow,$num);
                    } else {
                        $outRow--;
                    }
                }
                // Save sequence
                $seqData[$partner] = $seq;
                file_put_contents($seqFile, json_encode($seqData));
                // Save file with prefix in name and .xls
                $outDir = __DIR__ . '/converted';
                if (!is_dir($outDir)) mkdir($outDir,0755,true);
                $filename = "{$prefix}_{$origName}.xls";
                IOFactory::createWriter($tpl,'Xls')->save("$outDir/$filename");
                $resultHtml = "<div class=\"card result\">
  <svg onclick=\"window.location.reload()\" xmlns=\"http://www.w3.org/2000/svg\" width=\"16\" height=\"16\" fill=\"var(--accent)\" viewBox=\"0 0 16 16\" style=\"margin-right:0.5rem; vertical-align:middle; cursor:pointer;\"><path d=\"M8 3a5 5 0 1 1-4.546 2.914.5.5 0 1 0-.908-.418A6 6 0 1 0 14 8h-1a5 5 0 0 1-5-5z\"/><path d=\"M8 0a.5.5 0 0 1 .5.5v2.5H10a.5.5 0 0 1 0 1H7a.5.5 0 0 1-.5-.5V.5A.5.5 0 0 1 8 0z\"/></svg><strong>Fajl spreman:</strong> <a href=\"converted/$filename\" target=\"_blank\">$filename</a>
</div>";
            }
        }
    }
}

?><!DOCTYPE html>
<html lang="sr">
<head><meta charset="UTF-8"><title>Excel Konverter</title><style>:root{--bg:#fff;--fg:#343541;--accent:#10a37f;--card:#f7f7f8}body{background:var(--card);color:var(--fg);font-family:Inter,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}.container{background:var(--bg);padding:2rem;border-radius:.5rem;box-shadow:0 4px 20px rgba(0,0,0,.1);width:100%;max-width:400px}h2{text-align:center;margin-top:0}select,input[type=file],button{width:100%;padding:.75rem;margin:.5rem 0;border:1px solid #d1d5db;border-radius:.375rem;font-size:1rem}button{background:var(--accent);color:#fff;border:none;cursor:pointer}button:hover{opacity:.9}.card{margin-top:1rem;padding:1rem;background:var(--card);border-radius:.5rem;text-align:center}.result{display:flex;align-items:center;justify-content:center}.result svg{margin-right:.5rem;width:16px;cursor:pointer}.error{color:#e01e5a;text-align:center}</style></head>
<body>
<div class="container">
<h2>Konvertuj Excel</h2>
<form method="post" enctype="multipart/form-data">
<label>Partner:<select name="partner"><option>CS</option><option>MB</option><option>PM</option><option>SL</option><option>EP</option><option>SV</option></select></label>
<label>Izaberi fajl:<input type="file" name="excel_file" accept=".xls,.xlsx" required></label>
<button type="submit">Pokreni</button>
</form>
<?php echo $resultHtml;?>
<p style="text-align:center;font-size:.875rem;color:#6b7280;margin-top:1rem;">© Marko Mladenović, <?php echo date('Y');?>. Sva prava zadržana.</p>
</div>
</body>
</html>
