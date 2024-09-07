<?php
declare(strict_types=1);

namespace Bud\Excel;

use Closure;
use Generator;
use RuntimeException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use Bud\Excel\Exception\ExcelException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use Hyperf\HttpMessage\Stream\SwooleStream;
use Hyperf\HttpServer\Contract\RequestInterface;
use Hyperf\HttpServer\Contract\ResponseInterface;
use Hyperf\Validation\Contract\ValidatorFactoryInterface;

abstract class PHPExcel
{
    /**
     * 导入字段验证规则，用法同 hyperf/validation 手动创建验证器的用法
     * @var array
     */
    protected array $import_roles = [];

    /**
     * 自定义导入验证错误消息，用法同 hyperf/validation 手动创建验证器的用法
     * @var array
     */
    protected array $import_messages = [];

    /**
     * 导入导出字段配置
     * @var array
     */
    protected array $property = [];

    public function __construct(
        protected RequestInterface          $request,
        protected ResponseInterface         $response,
        protected ValidatorFactoryInterface $validationFactory
    )
    {
        if (empty($this->property)) throw new ExcelException('Field properties cannot be empty！');
    }

    /**
     * 解析导入的数据
     * @param \Closure|null $closure 匿名函数，自定义插入逻辑，建议传递。唯一参数当前解析的记录数组
     * @param string $fileKey 上传文件的检索键。默认：file
     * @return array|bool $closure为 null 时返回表格中的数据，为函数时返回 bool
     */
    public function parseImport(?\Closure $closure = null, string $fileKey = 'file'): array|bool
    {
        $data = [];
        if (!$this->request->hasFile($fileKey)) return false;
        $file = $this->request->file($fileKey);
        $tempFileName = 'import_' . time() . '.' . $file->getExtension();
        $tempFilePath = BASE_PATH . '/runtime/' . $tempFileName;
        file_put_contents($tempFilePath, $file->getStream()->getContents());
        $reader = IOFactory::createReader(IOFactory::identify($tempFilePath));
        $reader->setReadDataOnly(true);
        $sheet = $reader->load($tempFilePath);
        $property = $this->parseProperty();
        $endCell = $this->getColumnIndex(count($property));
        foreach ($sheet->getActiveSheet()->getRowIterator(2) as $row) {
            $temp = [];
            foreach ($row->getCellIterator('A', $endCell, true) as $index => $item) {
                $column = ord($index) - 65;
                // 存在非仅导出的列并且值不为空才赋值
                if (isset($property[$column]) && !$property[$column]['only_export'] && !empty($item->getFormattedValue())) {
                    $dict = $property[$column]['dictData'];
                    // 字典为空或者值存在于字典中则直接赋值
                    if (empty($dict) || isset($dict[$item->getFormattedValue()])) {
                        $temp[$property[$column]['name']] = $item->getFormattedValue();
                    } else {
                        // 否则反转字典键值进行再次匹配（比如：数据库中以0和1标识启用状态时，表格中录入时可以直接是0和1也可以是对应的字典值）
                        $dict = array_flip($dict);
                        $temp[$property[$column]['name']] = isset($dict[$item->getFormattedValue()]) ? $dict[$item->getFormattedValue()] : $item->getFormattedValue();
                    }
                }
            }
            if (!empty($temp)) {
                if (!empty($this->import_roles)) {
                    $validator = $this->validationFactory->make($temp, $this->import_roles, $this->import_messages);
                    if ($validator->fails()) {
                        unlink($tempFilePath);
                        throw new ExcelException("第{$row->getRowIndex()}行：" . $validator->errors()->first(), 422);
                    }
                }
                if ($closure instanceof \Closure) {
                    try {
                        $closure($temp);
                    } catch (\Throwable $exception) {
                        unlink($tempFilePath);
                        throw new ExcelException($exception->getMessage(), 500, $exception);
                    }
                    continue;
                }
                $data[] = $temp;
            }
        }
        unlink($tempFilePath);
        return $closure instanceof \Closure ? true : $data;
    }


    /**
     * 导出
     * @param string $filename 导出文件名
     * @param array|\Closure $closure 导出的数据。。匿名函数时必须返回数组
     * @param array $property 导出字段属性配置
     * @param string $format 导出格式：Xlsx|Xls|Csv 默认 Xlsx
     * @return \Psr\Http\Message\ResponseInterface
     */
    public function export(string $filename, array|Closure $closure, array $property = [], string $format = 'Xlsx'): \Psr\Http\Message\ResponseInterface
    {
        $spread = new Spreadsheet();
        $sheet = $spread->getActiveSheet();
        is_array($closure) ? $data = $closure : $data = $closure();
        !empty($property) && $this->property = $property;
        // 表头
        $titleStart = 0;
        foreach ($this->parseProperty() as $item) {
            $headerColumn = $this->getColumnIndex($titleStart) . '1';
            $sheet->setCellValue($headerColumn, $item['title']);
            $style = $sheet->getStyle($headerColumn)->getFont()->setBold(true);
            $columnDimension = $sheet->getColumnDimension($headerColumn[0]);
            empty($item['width']) ? $columnDimension->setAutoSize(true) : $columnDimension->setWidth((float)$item['width']);
            empty($item['align']) || $sheet->getStyle($headerColumn)->getAlignment()->setHorizontal($item['align']);
            empty($item['headColor']) || $style->setColor(new Color(str_replace('#', '', $item['headColor'])));
            if (!empty($item['headBgColor'])) {
                $sheet->getStyle($headerColumn)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB(str_replace('#', '', $item['headBgColor']));
            }
            $titleStart++;
        }
        $generate = $this->yieldExcelData($data);
        // 表体
        try {
            $row = 2;
            while ($generate->valid()) {
                $column = 0;
                $items = $generate->current();
                foreach ($items as $name => $value) {
                    $columnRow = $this->getColumnIndex($column) . $row;
                    $annotation = [];
                    foreach ($this->parseProperty() as $item) {
                        if ($item['name'] == $name) {
                            $annotation = $item;
                            break;
                        }
                    }
                    if (!empty($annotation['dictData']) && $annotation['dictData'][$value]) {
                        $sheet->setCellValue($columnRow, $annotation['dictData'][$value]);
                    } else {
                        $sheet->setCellValue($columnRow, is_string($value) ? $value : $value . "\t");
                    }
                    empty($annotation['align']) || $sheet->getStyle($columnRow)->getAlignment()->setHorizontal($annotation['align']);
                    empty($annotation['color']) || $sheet->getStyle($columnRow)->getFont()->setColor(new Color(str_replace('#', '', $annotation['color'])));
                    empty($annotation['bgColor']) || $sheet->getStyle($columnRow)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB(str_replace('#', '', $annotation['bgColor']));
                    $column++;
                }
                $generate->next();
                $row++;
            }
        } catch (RuntimeException) {
        }
        $format = in_array(ucwords($format), ['Xlsx', 'Xls', 'Csv']) ? ucwords($format) : 'Xlsx';
        $writer = IOFactory::createWriter($spread, $format);
        ob_start();
        $writer->save('php://output');
        $res = $this->downloadExcel($filename . '.' . strtolower($format), ob_get_contents());
        ob_end_clean();
        $spread->disconnectWorksheets();
        return $res;
    }

    /**
     * 构造导出数据迭代器
     * @param array $data
     * @return \Generator
     */
    protected function yieldExcelData(array $data): Generator
    {
        foreach ($data as $dat) {
            $yield = [];
            foreach ($this->parseProperty() as $item) {
                $yield[$item['name']] = $dat[$item['name']] ?? '';
            }
            yield $yield;
        }
    }

    /**
     * 解析字段配置
     * @return array
     */
    protected function parseProperty(): array
    {
        $data = [];
        foreach ($this->property as $name => $mate) {
            $data[$mate['index']] = [
                'name' => $name,
                'title' => $mate['title'] ?? $name,
                'width' => $mate['width'] ?? null,
                'align' => $mate['align'] ?? null,
                'headColor' => $mate['headColor'] ?? null,
                'headBgColor' => $mate['headBgColor'] ?? null,
                'color' => $mate['color'] ?? null,
                'bgColor' => $mate['bgColor'] ?? null,
                'only_export' => $mate['only_export'] ?? false,
                'dictData' => $mate['dictData'] ?? []
            ];
        }
        ksort($data);
        return $data;
    }

    /**
     * 获取 excel 列索引
     * @param int $columnIndex
     * @return string
     */
    protected function getColumnIndex(int $columnIndex = 0): string
    {
        if ($columnIndex < 26) {
            return chr(65 + $columnIndex);
        } else if ($columnIndex < 702) {
            return chr(64 + intval($columnIndex / 26)) . chr(65 + $columnIndex % 26);
        } else {
            return chr(64 + intval(($columnIndex - 26) / 676)) . chr(65 + intval((($columnIndex - 26) % 676) / 26)) . chr(65 + $columnIndex % 26);
        }
    }

    /**
     * 下载excel
     * @param string $filename
     * @param string $content
     * @return \Psr\Http\Message\ResponseInterface
     */
    protected function downloadExcel(string $filename, string $content): \Psr\Http\Message\ResponseInterface
    {
        return $this->response
            ->withHeader('content-description', 'File Transfer')
            ->withHeader('content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            ->withHeader('Content-Disposition', 'attachment; filename="' . $filename . '"')
            ->withHeader('content-transfer-encoding', 'binary')
            ->withHeader('pragma', 'public')
            ->withBody(new SwooleStream($content));
    }
}
