<?php
require_once __DIR__ . '/helpers.php';
require_login();

require_once __DIR__ . '/vendor/autoload.php';
require_once __DIR__ . '/includes/export_tokens.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

$ttlDays = max(1, min(365, (int)($_GET['ttl'] ?? 30)));
$qrSize  = max(120, min(300, (int)($_GET['qr'] ?? 160)));

$selectedIds = [];
if (!empty($_REQUEST['selected'])) {
    $selectedIds = array_filter(array_map('intval', explode(',', (string)$_REQUEST['selected'])));
}

if ($selectedIds) {
    $tasks = fetch_tasks_by_ids($selectedIds);
} else {
    $tasks = export_tasks(get_filter_values());
}

$taskIds = array_column($tasks, 'id');
$pdo     = get_pdo();

task_export_ensure_token_tables($pdo);
$existingTokens = task_export_fetch_tokens($pdo, $taskIds);
$baseUrl        = base_url_for_pdf();
$publicPath     = '/public_task_photos.php';

$qrMap = [];

foreach ($tasks as $task) {
    $taskId = (int)$task['id'];
    if ($taskId <= 0) {
        continue;
    }

    $tokenRow = $existingTokens[$taskId] ?? task_export_insert_token($pdo, $taskId, $ttlDays);
    $token    = is_string($tokenRow['token']) ? $tokenRow['token'] : (string)$tokenRow['token'];

    $url = $baseUrl . $publicPath . '?t=' . rawurlencode($token);
    $qr = qr_data_uri($url, $qrSize);
    if ($qr) {
        $qrMap[$taskId] = $qr;
    }
}

$spreadsheet = new Spreadsheet();
$sheet       = $spreadsheet->getActiveSheet();
$sheet->setTitle('Tasks');

$headers = [
    'ID', 'Building', 'Room', 'Title', 'Description',
    'Priority', 'Status', 'Assigned To',
    'Due Date', 'Created At', 'Updated At',
    'QR Code'
];
$sheet->fromArray($headers, null, 'A1');

$headerStyle = [
    'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
    'fill' => ['fillType' => 'solid', 'startColor' => ['rgb' => '2563EB']],
];
$sheet->getStyle('A1:L1')->applyFromArray($headerStyle);

$row       = 2;
$tempFiles = [];

foreach ($tasks as $task) {
    $taskId = (int)$task['id'];

    $sheet->fromArray([
        $task['id'],
        $task['building_name'],
        $task['room_number'] . ($task['room_label'] ? ' - ' . $task['room_label'] : ''),
        $task['title'],
        $task['description'],
        $task['priority'],
        $task['status'],
        $task['assigned_to'],
        $task['due_date'],
        $task['created_at'],
        $task['updated_at'],
        '',
    ], null, 'A' . $row);

    if (!empty($qrMap[$taskId])) {
        $parts = explode(',', $qrMap[$taskId], 2);
        $binary = isset($parts[1]) ? base64_decode($parts[1]) : base64_decode($parts[0]);
        if ($binary !== false) {
            $tmp = tempnam(sys_get_temp_dir(), 'qr_');
            if ($tmp && file_put_contents($tmp, $binary) !== false) {
                $drawing = new Drawing();
                $drawing->setName('QR ' . $taskId);
                $drawing->setPath($tmp);
                $drawing->setCoordinates('L' . $row);
                $drawing->setHeight($qrSize * 0.9);
                $drawing->setWorksheet($sheet);
                $sheet->getRowDimension($row)->setRowHeight(max($sheet->getRowDimension($row)->getRowHeight(), $qrSize * 0.9));
                $tempFiles[] = $tmp;
            } elseif ($tmp) {
                @unlink($tmp);
            }
        }
    }

    $row++;
}

$sheet->getStyle('A1:L' . ($row - 1))->applyFromArray([
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['rgb' => 'CCCCCC']
        ]
    ]
]);

foreach (range('A', 'K') as $col) {
    $sheet->getColumnDimension($col)->setAutoSize(true);
}
$sheet->getColumnDimension('L')->setWidth(18);

if ($row > 2) {
    for ($r = 2; $r < $row; $r++) {
        if ($sheet->getRowDimension($r)->getRowHeight() < 80) {
            $sheet->getRowDimension($r)->setRowHeight(80);
        }
    }
}

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="tasks_with_qr.xlsx"');
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');

foreach ($tempFiles as $file) {
    if (is_string($file) && file_exists($file)) {
        @unlink($file);
    }
}
exit;

function task_export_ensure_token_tables(PDO $pdo): void
{
    $pdo->exec(
        "CREATE TABLE IF NOT EXISTS public_task_tokens (
            id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            task_id BIGINT UNSIGNED NOT NULL,
            token VARBINARY(32) NOT NULL,
            expires_at DATETIME NOT NULL,
            revoked TINYINT(1) NOT NULL DEFAULT 0,
            use_count INT UNSIGNED NOT NULL DEFAULT 0,
            last_used_at DATETIME NULL,
            UNIQUE KEY uniq_token (token),
            INDEX idx_task_exp (task_id, expires_at),
            INDEX idx_exp (expires_at)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4"
    );

    $pdo->exec(
        "CREATE TABLE IF NOT EXISTS public_task_hits (
            id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            token_id BIGINT UNSIGNED NOT NULL,
            task_id BIGINT UNSIGNED NOT NULL,
            ts DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            ip VARBINARY(16) NULL,
            ua VARCHAR(255) NULL,
            INDEX idx_token (token_id),
            INDEX idx_task (task_id),
            CONSTRAINT fk_task_hits_token FOREIGN KEY (token_id)
              REFERENCES public_task_tokens(id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4"
    );
}

function task_export_fetch_tokens(PDO $pdo, array $taskIds): array
{
    if (!$taskIds) {
        return [];
    }

    $placeholders = implode(',', array_fill(0, count($taskIds), '?'));
    $sql = "SELECT t1.*
            FROM public_task_tokens t1
            JOIN (
                SELECT task_id, MAX(id) AS max_id
                FROM public_task_tokens
                WHERE revoked = 0 AND expires_at > NOW() AND task_id IN ($placeholders)
                GROUP BY task_id
            ) t2 ON t1.id = t2.max_id";
    $stmt = $pdo->prepare($sql);
    $stmt->execute($taskIds);

    $rows = $stmt->fetchAll(PDO::FETCH_ASSOC) ?: [];
    $tokens = [];
    foreach ($rows as $row) {
        $tokens[(int)$row['task_id']] = $row;
    }

    return $tokens;
}

function task_export_insert_token(PDO $pdo, int $taskId, int $ttlDays): array
{
    $token  = task_export_random_token();
    $expiry = (new DateTimeImmutable('now'))->modify("+{$ttlDays} days")->format('Y-m-d H:i:s');
    $stmt = $pdo->prepare('INSERT INTO public_task_tokens (task_id, token, expires_at) VALUES (:task, :token, :expires)');
    $stmt->execute([
        ':task'   => $taskId,
        ':token'  => $token,
        ':expires'=> $expiry,
    ]);

    $id = (int)$pdo->lastInsertId();
    return [
        'id'          => $id,
        'task_id'     => $taskId,
        'token'       => $token,
        'expires_at'  => $expiry,
        'revoked'     => 0,
        'use_count'   => 0,
        'last_used_at'=> null,
    ];
}

function task_export_random_token(): string
{
    return rtrim(strtr(base64_encode(random_bytes(16)), '+/', '-_'), '=');
}
