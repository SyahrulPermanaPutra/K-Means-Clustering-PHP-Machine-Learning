<?php

require "vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Phpml\Clustering\KMeans;

// ===============================
// LOAD DATASET
// ===============================
$file = __DIR__ . "/winequality-red.csv";

// Cek jika file exists
if (!file_exists($file)) {
    die("File $file tidak ditemukan!");
}

$spreadsheet = IOFactory::load($file);
$data = $spreadsheet->getActiveSheet()->toArray();

$header = array_shift($data); // remove header

$dataset = [];
$originalData = []; // Simpan data asli termasuk quality

foreach ($data as $row) {
    // Filter baris yang kosong atau tidak valid
    if (empty($row) || count(array_filter($row, 'strlen')) == 0) {
        continue;
    }
    
    // ambil semua kolom KECUALI QUALITY (kolom terakhir) untuk clustering
    $features = array_slice($row, 0, count($row) - 1);
    
    // Konversi ke float dan pastikan tidak ada nilai null
    $features = array_map(function($value) {
        return floatval($value);
    }, $features);
    
    $dataset[] = $features;
    $originalData[] = $row; // Simpan data lengkap
}

echo "Jumlah data yang dimuat: " . count($dataset) . "\n";
echo "Dimensi setiap data: " . (count($dataset) > 0 ? count($dataset[0]) : 0) . "\n";

// ===============================
// FUNCTION HITUNG WCSS MANUAL - VERSI AMAN
// ===============================
function getWCSS_safe($k, $data)
{
    if (empty($data) || $k < 1) {
        return 0;
    }

    try {
        $kmeans = new KMeans($k);
        $clusters = $kmeans->cluster($data);

        $totalWCSS = 0;

        foreach ($clusters as $clusterIndex => $cluster) {
            // Skip jika cluster kosong
            if (empty($cluster) || !is_array($cluster)) {
                continue;
            }

            // Cek elemen pertama untuk mendapatkan dimensi
            $firstPoint = null;
            foreach ($cluster as $point) {
                if (is_array($point) && !empty($point)) {
                    $firstPoint = $point;
                    break;
                }
            }

            if ($firstPoint === null) {
                continue; // Tidak ada data valid dalam cluster
            }

            $dimension = count($firstPoint);
            $centroid = array_fill(0, $dimension, 0);
            $pointCount = 0;

            // Hitung centroid
            foreach ($cluster as $point) {
                if (!is_array($point) || count($point) != $dimension) {
                    continue; // Skip point yang tidak valid
                }
                
                for ($i = 0; $i < $dimension; $i++) {
                    $centroid[$i] += $point[$i];
                }
                $pointCount++;
            }

            if ($pointCount == 0) {
                continue;
            }

            for ($i = 0; $i < $dimension; $i++) {
                $centroid[$i] /= $pointCount;
            }

            // Hitung WCSS
            foreach ($cluster as $point) {
                if (!is_array($point) || count($point) != $dimension) {
                    continue;
                }
                
                $dist = 0;
                for ($i = 0; $i < $dimension; $i++) {
                    $dist += pow($point[$i] - $centroid[$i], 2);
                }
                $totalWCSS += $dist;
            }
        }

        return $totalWCSS;
        
    } catch (Exception $e) {
        echo "ERROR dalam getWCSS_safe: " . $e->getMessage() . "\n";
        return null;
    }
}

// ===============================
// TENTUKAN NILAI K OPTIMAL DENGAN ELBOW METHOD
// ===============================
echo "\n=== PERHITUNGAN ELBOW METHOD ===\n";

$maxK = 8;
$wcssResults = [];
$bestK = 3; // default

for ($k = 1; $k <= $maxK; $k++) {
    echo "Menghitung K = $k...\n";
    
    $wcss = getWCSS_safe($k, $dataset);
    $wcssResults[$k] = $wcss;
    echo "K = $k → WCSS = " . ($wcss ?? 'ERROR') . "\n";
}

// Cari elbow point sederhana
if (count($wcssResults) > 2) {
    $reductions = [];
    for ($k = 2; $k <= $maxK; $k++) {
        if (isset($wcssResults[$k-1]) && isset($wcssResults[$k]) && $wcssResults[$k-1] && $wcssResults[$k]) {
            $reduction = $wcssResults[$k-1] - $wcssResults[$k];
            $reductions[$k] = $reduction;
        }
    }
    
    if (!empty($reductions)) {
        $avgReduction = array_sum($reductions) / count($reductions);
        foreach ($reductions as $k => $reduction) {
            if ($reduction < $avgReduction * 0.5) {
                $bestK = $k - 1;
                break;
            }
        }
    }
}

echo "K optimal yang dipilih: $bestK\n";

// ===============================
// CLUSTERING DENGAN K OPTIMAL - VERSI YANG LEBIH AMAN
// ===============================
echo "\n=== MELAKUKAN CLUSTERING DENGAN K = $bestK ===\n";

try {
    $kmeans = new KMeans($bestK);
    $clusters = $kmeans->cluster($dataset);
    
    // Mapping cluster assignments
    $clusterAssignments = [];
    $currentIndex = 0;
    
    foreach ($clusters as $clusterIndex => $cluster) {
        if (is_array($cluster)) {
            $clusterSize = count($cluster);
            for ($i = 0; $i < $clusterSize; $i++) {
                $clusterAssignments[$currentIndex] = $clusterIndex + 1;
                $currentIndex++;
            }
        }
    }
    
    // Jika mapping tidak sesuai, gunakan metode alternatif
    if (count($clusterAssignments) !== count($dataset)) {
        echo "Menggunakan metode alternatif untuk mapping cluster...\n";
        $clusterAssignments = [];
        
        // Hitung centroid untuk setiap cluster
        $centroids = [];
        foreach ($clusters as $clusterIndex => $cluster) {
            if (empty($cluster) || !is_array($cluster)) {
                continue;
            }
            
            $firstPoint = null;
            foreach ($cluster as $point) {
                if (is_array($point) && !empty($point)) {
                    $firstPoint = $point;
                    break;
                }
            }
            
            if ($firstPoint === null) {
                continue;
            }
            
            $dimension = count($firstPoint);
            $centroid = array_fill(0, $dimension, 0);
            $pointCount = 0;
            
            foreach ($cluster as $point) {
                if (!is_array($point) || count($point) != $dimension) {
                    continue;
                }
                
                for ($i = 0; $i < $dimension; $i++) {
                    $centroid[$i] += $point[$i];
                }
                $pointCount++;
            }
            
            if ($pointCount > 0) {
                for ($i = 0; $i < $dimension; $i++) {
                    $centroid[$i] /= $pointCount;
                }
                $centroids[$clusterIndex] = $centroid;
            }
        }
        
        // Assign setiap data point ke cluster terdekat
        foreach ($dataset as $dataIndex => $point) {
            $minDistance = PHP_FLOAT_MAX;
            $assignedCluster = 1;
            
            foreach ($centroids as $clusterIndex => $centroid) {
                $distance = 0;
                for ($i = 0; $i < count($point); $i++) {
                    $distance += pow($point[$i] - $centroid[$i], 2);
                }
                $distance = sqrt($distance);
                
                if ($distance < $minDistance) {
                    $minDistance = $distance;
                    $assignedCluster = $clusterIndex + 1;
                }
            }
            
            $clusterAssignments[$dataIndex] = $assignedCluster;
        }
    }
    
    // ===============================
    // HITUNG STATISTIK UNTUK SHEET BARU
    // ===============================
    $totalData = count($dataset);
    $clusterCounts = array_count_values($clusterAssignments);
    
    // Siapkan data untuk sheet statistik detail
    $statistikDetail = [];
    for ($i = 1; $i <= $bestK; $i++) {
        $clusterSize = isset($clusterCounts[$i]) ? $clusterCounts[$i] : 0;
        $percentage = $totalData > 0 ? ($clusterSize / $totalData) * 100 : 0;
        
        $statistikDetail[] = [
            'Cluster ' . $i,
            $clusterSize,
            number_format($percentage, 2) . '%',
            $percentage >= 50 ? 'MAJORITY' : ($percentage >= 30 ? 'MEDIUM' : 'MINORITY')
        ];
    }
    
    // Tambahkan total
    $statistikDetail[] = [
        'TOTAL',
        $totalData,
        '100%',
        'ALL CLUSTERS'
    ];
    
    // ===============================
    // EKSPOR HASIL KE EXCEL
    // ===============================
    $outputSpreadsheet = new Spreadsheet();
    
    // ===============================
    // SHEET 1: HASIL CLUSTERING
    // ===============================
    $outputSheet = $outputSpreadsheet->getActiveSheet();
    $outputSheet->setTitle('Hasil Clustering');
    
    // Prepare header
    $outputHeader = array_merge($header, ['Cluster']);
    $outputSheet->fromArray([$outputHeader], null, 'A1');
    
    // Style untuk header - DIPERBAIKI
    $headerStyle = [
        'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID, 
            'startColor' => ['rgb' => '4472C4']
        ]
    ];
    $outputSheet->getStyle('A1:' . chr(65 + count($outputHeader) - 1) . '1')->applyFromArray($headerStyle);
    
    // Tambahkan data ke Excel
    $rowIndex = 2;
    foreach ($originalData as $dataIndex => $originalRow) {
        $clusterNumber = isset($clusterAssignments[$dataIndex]) ? $clusterAssignments[$dataIndex] : 0;
        $outputRow = array_merge($originalRow, [$clusterNumber]);
        $outputSheet->fromArray([$outputRow], null, 'A' . $rowIndex);
        $rowIndex++;
    }
    
    // Auto size columns
    foreach (range('A', chr(65 + count($outputHeader) - 1)) as $col) {
        $outputSheet->getColumnDimension($col)->setAutoSize(true);
    }
    
    // ===============================
    // SHEET 2: STATISTIK DETAIL CLUSTER - DIPERBAIKI
    // ===============================
    $statistikSheet = $outputSpreadsheet->createSheet();
    $statistikSheet->setTitle('Statistik Detail');
    
    // Header untuk statistik detail
    $statistikHeaders = ['Cluster', 'Jumlah Data', 'Persentase', 'Kategori'];
    $statistikSheet->fromArray([$statistikHeaders], null, 'A1');
    $statistikSheet->getStyle('A1:D1')->applyFromArray($headerStyle);
    
    // Tambahkan data statistik
    $statistikSheet->fromArray($statistikDetail, null, 'A2');
    
    // Style conditional untuk kategori - DIPERBAIKI
    $lastStatRow = 1 + count($statistikDetail);
    for ($row = 2; $row <= $lastStatRow; $row++) {
        $kategoriCell = 'D' . $row;
        $kategoriValue = $statistikSheet->getCell($kategoriCell)->getValue();
        
        $color = 'FFFFFF'; // default white
        
        if ($kategoriValue === 'MAJORITY') {
            $color = 'C6EFCE'; // hijau muda
        } elseif ($kategoriValue === 'MEDIUM') {
            $color = 'FFEB9C'; // kuning muda
        } elseif ($kategoriValue === 'MINORITY') {
            $color = 'FFC7CE'; // merah muda
        } elseif ($kategoriValue === 'ALL CLUSTERS') {
            $color = 'D9E1F2'; // biru muda
        }
        
        // PERBAIKAN: Gunakan applyFromArray untuk styling
        $statistikSheet->getStyle($kategoriCell)
            ->applyFromArray([
                'fill' => [
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'startColor' => ['rgb' => $color]
                ]
            ]);
    }
    
    // Tambahkan chart description
    $statistikSheet->setCellValue('F1', 'ANALISIS DISTRIBUSI CLUSTER');
    $statistikSheet->getStyle('F1')->getFont()->setBold(true)->setSize(14);
    
    $statistikSheet->setCellValue('F3', 'Keterangan:');
    $statistikSheet->setCellValue('F4', '• MAJORITY: ≥ 50% data');
    $statistikSheet->setCellValue('F5', '• MEDIUM: 30% - 49.9% data');
    $statistikSheet->setCellValue('F6', '• MINORITY: < 30% data');
    
    // Style untuk keterangan
    $statistikSheet->getStyle('F3:F6')->getFont()->setBold(true);
    
    // Auto size columns
    foreach (range('A', 'F') as $col) {
        $statistikSheet->getColumnDimension($col)->setAutoSize(true);
    }
    
    // ===============================
    // SHEET 3: SUMMARY CLUSTER
    // ===============================
    $summarySheet = $outputSpreadsheet->createSheet();
    $summarySheet->setTitle('Summary Cluster');
    
    // Header summary
    $summaryHeaders = ['Cluster', 'Jumlah Data', 'Persentase'];
    $summarySheet->fromArray([$summaryHeaders], null, 'A1');
    $summarySheet->getStyle('A1:C1')->applyFromArray($headerStyle);
    
    // Data summary
    $summaryData = [];
    for ($i = 1; $i <= $bestK; $i++) {
        $clusterSize = isset($clusterCounts[$i]) ? $clusterCounts[$i] : 0;
        $percentage = $totalData > 0 ? ($clusterSize / $totalData) * 100 : 0;
        
        $summaryData[] = [
            'Cluster ' . $i,
            $clusterSize,
            number_format($percentage, 2) . '%'
        ];
    }
    
    // Tambahkan total
    $summaryData[] = [
        'TOTAL',
        $totalData,
        '100%'
    ];
    
    $summarySheet->fromArray($summaryData, null, 'A2');
    
    // Auto size columns
    foreach (range('A', 'C') as $col) {
        $summarySheet->getColumnDimension($col)->setAutoSize(true);
    }
    
    // ===============================
    // SHEET 4: CENTROID CLUSTER
    // ===============================
    $centroidSheet = $outputSpreadsheet->createSheet();
    $centroidSheet->setTitle('Centroid Cluster');
    
    // Header centroid (fitur tanpa quality)
    $featureHeaders = array_slice($header, 0, count($header) - 1);
    $centroidHeaders = array_merge(['Cluster'], $featureHeaders);
    $centroidSheet->fromArray([$centroidHeaders], null, 'A1');
    $centroidSheet->getStyle('A1:' . chr(65 + count($centroidHeaders) - 1) . '1')->applyFromArray($headerStyle);
    
    // Hitung centroid untuk setiap cluster berdasarkan assignment
    $centroidData = [];
    for ($clusterNum = 1; $clusterNum <= $bestK; $clusterNum++) {
        $clusterPoints = [];
        
        // Kumpulkan semua point yang termasuk dalam cluster ini
        foreach ($dataset as $dataIndex => $point) {
            if (isset($clusterAssignments[$dataIndex]) && $clusterAssignments[$dataIndex] == $clusterNum) {
                $clusterPoints[] = $point;
            }
        }
        
        if (empty($clusterPoints)) {
            continue;
        }
        
        $dimension = count($clusterPoints[0]);
        $centroid = array_fill(0, $dimension, 0);
        
        foreach ($clusterPoints as $point) {
            for ($i = 0; $i < $dimension; $i++) {
                $centroid[$i] += $point[$i];
            }
        }
        
        for ($i = 0; $i < $dimension; $i++) {
            $centroid[$i] /= count($clusterPoints);
        }
        
        $centroidRow = array_merge(['Cluster ' . $clusterNum], $centroid);
        $centroidData[] = $centroidRow;
    }
    
    if (!empty($centroidData)) {
        $centroidSheet->fromArray($centroidData, null, 'A2');
        
        // Format angka centroid
        $lastCentroidRow = 1 + count($centroidData);
        $centroidSheet->getStyle('B2:' . chr(65 + count($centroidHeaders) - 1) . $lastCentroidRow)
                      ->getNumberFormat()
                      ->setFormatCode('0.0000');
    }
    
    // Auto size columns
    foreach (range('A', chr(65 + count($centroidHeaders) - 1)) as $col) {
        $centroidSheet->getColumnDimension($col)->setAutoSize(true);
    }
    
    // Set sheet pertama sebagai aktif
    $outputSpreadsheet->setActiveSheetIndex(0);
    
    // ===============================
    // SIMPAN FILE EXCEL
    // ===============================
    $timestamp = date('Y-m-d_H-i-s');
    $outputFilename = __DIR__ . "/clustering_results_{$bestK}K_{$timestamp}.xlsx";
    
    $writer = new Xlsx($outputSpreadsheet);
    $writer->save($outputFilename);
    
    echo "\n✅ HASIL CLUSTERING BERHASIL DIEXPORT KE EXCEL!\n";
    echo "📁 File: " . basename($outputFilename) . "\n";
    echo "📊 Jumlah Cluster: $bestK\n";
    echo "📈 Total Data: " . count($dataset) . "\n";
    
    // Tampilkan preview hasil
    echo "\n=== PREVIEW HASIL CLUSTERING ===\n";
    for ($i = 1; $i <= $bestK; $i++) {
        $clusterSize = isset($clusterCounts[$i]) ? $clusterCounts[$i] : 0;
        $percentage = number_format(($clusterSize / $totalData) * 100, 2);
        echo "Cluster " . $i . ": $clusterSize data ($percentage%)\n";
    }
    
    echo "\n📋 Sheet yang dihasilkan:\n";
    echo "1. Hasil Clustering - Data lengkap dengan label cluster\n";
    echo "2. Statistik Detail - Tabel dengan kategori (MAJORITY/MEDIUM/MINORITY)\n";
    echo "3. Summary Cluster - Ringkasan jumlah data per cluster\n";
    echo "4. Centroid Cluster - Nilai centroid setiap cluster\n";
    
} catch (Exception $e) {
    echo "❌ ERROR: " . $e->getMessage() . "\n";
    echo "Stack trace: " . $e->getTraceAsString() . "\n";
}

// ===============================
// TAMPILKAN HASIL WCSS
// ===============================
echo "\n=== HASIL PERHITUNGAN WCSS ===\n";
foreach ($wcssResults as $k => $wcss) {
    $indicator = ($k == $bestK) ? " ← OPTIMAL" : "";
    echo "K = $k: " . ($wcss ?? 'ERROR') . "$indicator\n";
}