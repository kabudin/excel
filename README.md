# rate-limit

> 适配 hyperf 框架的后端Excel导入导出组件，基于 PhpOffice\PhpSpreadsheet 实现

## 安装

```shell
composer require bud/excel
```

## 使用方法

```php
<?php
declare(strict_types=1);

namespace App\Controller;

use App\Service\UserService;
use Hyperf\Di\Annotation\Inject;
use Psr\Http\Message\ResponseInterface;
use Hyperf\HttpServer\Contract\RequestInterface;
use Hyperf\HttpServer\Annotation\AutoController;

#[AutoController(prefix: "test")]
class TestController extends AbstractController
{
    #[Inject]
    protected UserService $service;
    
    #[RequestMapping(path: "export", methods: "post")]
    public function export(): ResponseInterface
    {
        // 模拟数据
        $data = [
            ['id' => 1, 'account' => 'kabuding', 'phone' => '18888888888', 'email' => 'kabuding@qq.com', 'nickname' => '卡布丁', 'status' => 1, 'created_at' => '2024-03-26 22:39:37'],
            ['id' => 2, 'account' => 'juyokeji', 'phone' => '15555555555', 'email' => 'nilargs@qq.com', 'nickname' => 'nilargs', 'status' => 1, 'created_at' => '2024-03-27 22:39:37'],
            ['id' => 3, 'account' => 'zmboy', 'phone' => '13688888888', 'email' => 'zmboy@qq.com', 'nickname' => '追梦男孩', 'status' => 0, 'created_at' => '2024-03-28 22:39:37'],
        ];
        $property = $this->request->post('property',[]);  // 前端自定义导出字段配置
        $format = $this->request->post('format','Xlsx');  // 前端自定义导出格式，支持：Xlsx|Xls|Csv 默认 Xlsx
        return $this->service->export('用户列表', $data, $property, $format);
    }

    #[RequestMapping(path: "import", methods: "post")]
    public function import()
    {
        return $this->service->import();
    }
}
```

```php
<?php
declare(strict_types=1);

namespace App\Service;

use Bud\Excel\PHPExcel;

class UserService extends PHPExcel
{
    /**
     * 导入字段验证规则
     * @var array|string[] 
     */
    protected array $import_roles = [
        'account' => 'required',
        'phone' => 'required',
        'email' => 'required|email',
        'nickname' => 'required'
    ];
    /**
     * 自定义错误描述
     * @var array|string[] 
     */
    protected array $import_messages = [
        'account.required' => '账号不可为空',
        'phone.required' => '手机号不可为空',
        'email.required' => '邮箱不可为空',
        'email.email' => '邮箱地址不合法',
        'nickname.required' => '昵称不可为空'
    ];

    /**
     * 导入导出字段属性配置
     * title：字段头部标题
     * index：列顺序
     * align：对齐方式 left|center|right 默认left
     * headColor：列表头字体颜色
     * headBgColor：列表头背景颜色
     * color：列表体字体颜色
     * bgColor：列表体背景颜色
     * width：列宽
     * dictData：字典键值数组
     * only_export：是否仅导出，默认false。导入时会忽略仅导出的字段
     * @var array|string[] 
     */
    protected array $property = [
        'account' => ['title' => '账号', 'index' => 0, 'align' => 'center'],
        'phone' => ['title' => '手机号', 'index' => 2, 'headBgColor' => 'red'],
        'email' => ['title' => '邮箱', 'index' => 3, 'color' => 'green'],
        'nickname' => ['title' => '昵称', 'index' => 1, 'headBgColor' => 'yellow'],
        'status' => ['title' => '状态', 'index' => 4, 'width' => 15, 'bgColor' => '#1cda45', 'dictData' => [0 => '启用', 1 => '禁用']],
        'created_at' => ['title' => '创建时间', 'index' => 5, 'only_export' => true]
    ];

    public function import(): bool
    {
        /**
         * 该方法接收两个可选参数，匿名函数：导入的每一行数据集均会执行一次回调。上传文件检索标识，默认为：'file'
         * 不传参时，方法返回一个导入的所有数据的二维数组列表。导入数据量比较大时会产生较大的内存开销，建议匿名函数逐条处理返回bool
         */
        return $this->parseImport(function ($item) {
            // 保存数据
            var_dump($item);
        }, 'file');
    }
}
```
