# **前言**

>作者简介：Quixn，专注于 Node.js 技术栈分享，前端从 JavaScript 到 Node.js,再到后端数据库，优质文章推荐，【小Q全栈指南】作者，Github 博客开源项目 [github.com/Quixn…](https://github.com/Quixn/miniQBlog")

大家好，我是Quixn。最近在做业务系统时，遇到到一个需求。需要系统导出word文档，内容是比较复杂的表格。

![](https://cdn.jsdelivr.net/gh/Quixn/image-hosting@main/src/exportWordDemo.jpg)

一般业务系统遇到这种需求，都是后端同学来做导出，前端同学只需接收文件即可。毕竟后端有很多这方面成熟的库可以使用，而前端则很少有这方面开源的插件可以使用。没办法，跟笔者对接的后端同学恰好是个没什么经验的，不会做word文档导出。

![](https://img.soogif.com/iKKlFXy5dHPFYw2ragAOuXYmF75lf1Ol.png?scope=mdnice)

既然如此，那就只能自己上了。Google了一圈之后，发现前端能使用的也就一个[docxtemplater](https://github.com/open-xml-templating/docxtemplater "docxtemplater")库,star数`2.1k`，使用起来非常便携，也十分满足此次的需求。大家有其他更好的库推荐，欢迎留言告知。

## 1 docxtemplater 简介

- 官网 https://docxtemplater.com/

- GitHub https://github.com/open-xml-templating/docxtemplater

`docxtemplater` 使用 JSON 数据格式作为输入，可以处理`docx` 和 `ppt`模板。不像一些其它的工具，比如 `docx.js, docx4j, python-docx` 等，需要自己编写代码来生成文件，`docxtemplater`只需要用户通过标签的形式编写模板，就可以生成文件。

## 2 使用教程

### 2.1 安装依赖

项目所需依赖：`docxtemplater`，`jszip`，`jszip-utils`，`pizzip`，`file-saver`。

>1、docxtemplater：这个插件可以通过预先写好的word，excel等文件模板生成对应带数据的文件
>
>2、pizzip：这个插件用来创建，读取或编辑.zip的文件（同步的，还有一个插件是jszip，异步的）
>
>3、jszip-utils：与jszip/pizzip一起使用，jszip-utils 提供一个getBinaryContent(path, data)接口，path即是文件的路径，支持AJAX get请求，data为读取的文件内容。
>
>4、file-saver：适合在客户端生成文件的工具，它提供的接口saveAs(blob, "1.docx")将会使用到，方便我们保存文件。
>
>5、docxtemplater-image-module-free：需要导出图片的话需要这个插件

npm 安装如下：

```js
npm install  docxtemplater pizzip --save  // 处理docx模板
npm install  jszip-utils --save
npm install  jszip --save   
npm install  file-saver --save  // 处理输出文件
```

### 2.2 创建word模板文件

- 创建word模板：public/test.docx

>vue cli3/vue cli4 在 `public` 文件下存放word模板test.docx;
>vue cli2 在`static`文件下存放word模板test.docx;

![](https://cdn.jsdelivr.net/gh/Quixn/image-hosting@main/src/sourceTree.jpg)

如果直接在代码编辑器内通过新建文件的方式创建`test.docx`后面会报错，应该和文件编码格式有关，所以需要进入项目文件夹内右键新建`docx`文件，`test.docx`内编辑后编辑器内可以看到`pulic`文件下多了一个`~$test.docx`文件；出现这个文件夹基本就可以了。

### 2.3 docxtemplater 语法

>{%img} 图片
>
>{#list}{/list} 循环、if判断
>
>{#list}{/list}{^list}{/list} if else 
>
>{str} 文字

`docxtemplater` 通过标签的形式来填充数据的，简单的数据我们可以使用`{} + 变量名`来实现简单的文本替换。


例如：

![](https://cdn.jsdelivr.net/gh/Quixn/image-hosting@main/src/t_m_d1.jpg)

复杂的数据，例如需要多选打√的，就需要使用`docxtemplater` 的条件判断语法来实现。

![](https://cdn.jsdelivr.net/gh/Quixn/image-hosting@main/src/2022-05-05_10-49-40.jpg)

实现如下：传入的变量`category1` 为`true`时，才会渲染打`√`的效果。此时要传入另一个变量`category_1`,值为`category1`取反。

![](https://cdn.jsdelivr.net/gh/Quixn/image-hosting@main/src/2022-05-05_10-52-30.jpg)

如果整个word文档里有很多这种需要打钩的需求，那么这种实现方式就有一个很大的`弊端`：需要定义很多的变量来控制是否显示打`√`。也可以使用`if-else`的语法实现。笔者也在网上看到一个另外的[解决方案](https://www.freesion.com/article/6149783847/ 'docxtemplater 导出word文档勾选框的默认勾选')，笔者试了过效果并不理想，无法实现需求，大家也可以尝试一下。大家如果有更好的解决方法，欢迎留言告知。

表格需求实现如下：

![](https://cdn.jsdelivr.net/gh/Quixn/image-hosting@main/src/2022-05-05_11-17-11.jpg)


### 2.4 docxtemplater 完整代码实现

创建一个`exportWordDocx.js`文件，定义一个`exportWordDocx`函数，接收三个入参，`demoUrl`代表导出的`word`文档模板路径，`docxData`代表模板文档里定义的`dept`、`applyDate` 等字段整合给`docxData`传入即可。`fileName`代表导出的文件名，方面重命名等操作。

将`demoUrl`传入给`JSZipUtils.getBinaryContent`方法读取模板文件的二进制内容，之后创建一个`PizZip`实例，内容为模板的内容,再创建并加载`docxtemplater`实例对象。

使用`doc.setData`方法设置模板变量的值，对象的键需要和模板上的变量名一致，值就是你要放在模板上的值。

这里有一个地方需要注意的是：如果你的定义放在模板上的值为`null`或者`undefined`，最后导出来的`word`文档里，相对应的地方会直接显示`undefined`。解决方法：`doc.setOptions` 方法里的`nullGetter`值返回设置为空即可。

最后，通过`saveAs`方法导出`Word`文档。


```js
import JSZipUtils from "jszip-utils";
import docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";
import PizZip from "pizzip";

export const exportWordDocx = (demoUrl, docxData, fileName) => {
    // 读取并获得模板文件的二进制内容
    JSZipUtils.getBinaryContent(
        demoUrl,
        function (error, content) {
            // 抛出异常
            if (error) {
                throw error;
            }

            // 创建一个PizZip实例，内容为模板的内容
            let zip = new PizZip(content);
            // 创建并加载docxtemplater实例对象
            let doc = new docxtemplater().loadZip(zip);
            // 去除未定义值所显示的undefined
            doc.setOptions({
                nullGetter: function () {
                    return "";
                }
            }); // 设置角度解析器
            // 设置模板变量的值，对象的键需要和模板上的变量名一致，值就是你要放在模板上的值

            doc.setData({
                ...docxData,
            });

            try {
                // 用模板变量的值替换所有模板变量
                doc.render();
            } catch (error) {
                // 抛出异常
                let e = {
                    message: error.message,
                    name: error.name,
                    stack: error.stack,
                    properties: error.properties,
                };
                console.log(JSON.stringify({ error: e }));
                throw error;
            }

            // 生成一个代表docxtemplater对象的zip文件（不是一个真实的文件，而是在内存中的表示）
            let out = doc.getZip().generate({
                type: "blob",
                mimeType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });
            // 将目标文件对象保存为目标类型的文件，并命名
            saveAs(out, fileName);
        }
    );
}

```

## 3 总结

至此，我们使用`docxtemplater`导出`word`文档的实践就告一段落了。文章从介绍前端使用`docxtemplater`导出`word`文档场景，再到简要介绍`docxtemplater`基本使用语法，再到完整的代码示例实现导出功能。

欢迎关注，公众号回复【`docxtemplater最接实践`】获取文章的全部源码。

**关于我 & Node交流群**

>大家好，我是 Quixn，专注于 Node.js 技术栈分享，前端从 JavaScript 到 Node.js,再到后端数据库，优质文章推荐。如果你对 Node.js 学习感兴趣的话（后续有计划也可以)，可以关注我，加我微信【 Quixn1314 】，拉你进交流群一起交流、学习、共建，或者关注我的公众号【 小Q全栈指南 】。Github 博客开源项目 [github.com/Quixn…](https://github.com/Quixn/miniQBlog")


欢迎加我微信【 Quixn1314 】，拉你 进 Node.js 高级进阶群，一起学Node，长期交流学习...




















