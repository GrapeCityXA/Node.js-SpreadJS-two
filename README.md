# Node.js运行SpreadJS处理 Excel方案讨论（二）

本专题上一篇博客介绍了怎样在nodejs环境中搭建SpreadJS的运行环境，本篇博客是此系列的第二篇，重点在于对比GcExcel for Java和Node.js中运行SpreadJS的性能比较。由于SpreadJS和GcExcel的功能非常丰富，本文仅选择最为常见的两个功能点做对比，分别是设置区域数据，并导出Excel文档。对 GcExcel 不熟悉的同学，可以先看一下上边链接的官网主页，如果用过Apache POI的小伙伴，看一下这张图就知道，相对于 SpreadJS with Node.js，GcExcel的确是个强大的对手：

![输入图片说明](https://gcdn.grapecity.com.cn/data/attachment/forum/202102/28/021139isaskxqfdq7qxsa3.png)


本次测试的几个大前提：
首先，众所周知，Node.js是基于V8引擎来执行JavaScript的，因此它的js也是基于事件机制的非阻塞单线程运行，文件的I/O都是异步执行的。单线程的好处在于编码简单，开发难度低，心智消耗相对较小（对咱们码农比较友好）；而且它的文件的I/O是异步执行的，不像Java这种需要创建、回收线程（Node.js的IO操作在底层也是线程，这里不做深入讨论），这方面开销较小。
但是，相对的，单线程在做复杂运算方面相比多线程语言，没有优势，无法利用多线程来有效调配多核CPU进行优化；对于Node.js来说这个问题并非无解，但咱们这个专题讨论的大前提是基于SpreadJS的，这方面我们并没有提供针对Node.js计算的优化包（这跟SpreadJS产品定位实在是相去甚远，且我们也有了更好的解决方案GcExcel），因此在Node.js中运行SpreadJS就只能是单线程JS。
因此，基于以上前提，为了“公平”起见，本篇中设计的测试案例，在两个环境（Java 和 Node.js）中，都采用单线程执行，并且选择了对Node.js更加有优势的批量I/O操作，看看在性能上能否有得一战。


GCExcel 测试代码和结果：
代码非常简单，一个Performance类中执行了1000次设置数据、导出Excel文档的操作。


```
public class Performance {

        public static void main(String[] args) {
                System.out.println(System.getProperty("user.dir") + "/sources/jsonData");
                String jsonStr = readTxtFileIntoStringArrList(System.getProperty("user.dir") + "/sources/jsonData");
                JSONArray jsonArr = JSON.parseArray(jsonStr);
                //JSONObject jsonObj = (JSONObject) jsonArr.get(0);
                //System.out.println(jsonObj.get("Film"));
                run(1000, jsonArr);
        }

        public static void run(int times, JSONArray dataArr) {
                String path = System.getProperty("user.dir") + "/results/";
                System.out.println(path + "result.xlsx");
                long start = new Date().getTime();
                for (int i = 0; i < times; i++) {
                        Workbook workbook = new Workbook();
                        IWorksheet worksheet = workbook.getWorksheets().get(0); 
                        for (int j = 0; j < dataArr.size(); j++) {
                                JSONObject jsonObj = (JSONObject) dataArr.get(j);
                                worksheet.getRange(j, 0, 1, 8).get(0).setValue(jsonObj.get("Film"));
                                worksheet.getRange(j, 0, 1, 8).get(1).setValue(jsonObj.get("Genre"));
                                worksheet.getRange(j, 0, 1, 8).get(2).setValue(jsonObj.get("Lead Studio"));
                                worksheet.getRange(j, 0, 1, 8).get(3).setValue(jsonObj.get("Audience Score %"));
                                worksheet.getRange(j, 0, 1, 8).get(4).setValue(jsonObj.get("Profitability"));
                                worksheet.getRange(j, 0, 1, 8).get(5).setValue(jsonObj.get("Rating"));
                                worksheet.getRange(j, 0, 1, 8).get(6).setValue(jsonObj.get("Worldwide Gross"));
                                worksheet.getRange(j, 0, 1, 8).get(7).setValue(jsonObj.get("Year"));
                        }
                        workbook.save(path + "result" + i + ".xlsx");
                }
                System.out.println("运行"+times+"次花费时常（ms）: " + (new Date().getTime() - start));

        }

        public static String readTxtFileIntoStringArrList(String filePath) {
                StringBuilder list = new StringBuilder();
                try {
                        String encoding = "GBK";
                        File file = new File(filePath);
                        if (file.isFile() && file.exists()) {
                                InputStreamReader read = new InputStreamReader(new FileInputStream(file), encoding);// 考虑到编码格式
                                BufferedReader bufferedReader = new BufferedReader(read);
                                String lineTxt = null;

                                while ((lineTxt = bufferedReader.readLine()) != null) {
                                        list.append(lineTxt);
                                }
                                bufferedReader.close();
                                read.close();
                        } else {
                                System.out.println("找不到指定的文件");
                        }
                } catch (Exception e) {
                        System.out.println("读取文件内容出错");
                        e.printStackTrace();
                }
                return list.toString();
        }

}
```

完整的工程zip请参考附件：GcExcelPerformanceSample.zip
运行方式：导入Eclipse后直接run as Application
运行结果：
![输入图片说明](https://gcdn.grapecity.com.cn/data/attachment/forum/202102/28/022310ds3m931oms5kyo29.png)


Node.js 与 SpreadJS的测试代码和结果：
同样，代码没什么好讲的，如果有问题，走传送门回到第一篇复习一下~

```
const fs = require('fs');

// Initialize the mock browser variables
const mockBrowser = require('mock-browser').mocks.MockBrowser;
global.window = mockBrowser.createWindow();
global.document = window.document;
global.navigator = window.navigator;
global.HTMLCollection = window.HTMLCollection;
global.getComputedStyle = window.getComputedStyle;

const fileReader = require('filereader');
global.FileReader = fileReader;

const GC = require('@grapecity/spread-sheets');
const GCExcel = require('@grapecity/spread-excelio');

GC.Spread.Sheets.LicenseKey = GCExcel.LicenseKey = "Your License";

const dataSource = require('./data');

function runPerformance(times) {

  const timer = `test in ${times} times`;
  console.time(timer);

  for(let t=0; t<times; t++) {
    // const hostDiv = document.createElement('div');
    // hostDiv.id = 'ss';
    // document.body.appendChild(hostDiv);
    const wb = new GC.Spread.Sheets.Workbook()//global.document.getElementById('ss'));
    const sheet = wb.getSheet(0);
    for(let i=0; i<dataSource.length; i++) {
      sheet.setValue(i, 0, dataSource[i]["Film"]);
      sheet.setValue(i, 1, dataSource[i]["Genre"]);
      sheet.setValue(i, 2, dataSource[i]["Lead Studio"]);
      sheet.setValue(i, 3, dataSource[i]["Audience Score %"]);
      sheet.setValue(i, 4, dataSource[i]["Profitability"]);
      sheet.setValue(i, 5, dataSource[i]["Rating"]);
      sheet.setValue(i, 6, dataSource[i]["Worldwide Gross"]);
      sheet.setValue(i, 7, dataSource[i]["Year"]);
    }
    exportExcelFile(wb, times, t);
  }
  
}

function exportExcelFile(wb, times, t) {
    const excelIO = new GCExcel.IO();
    excelIO.save(wb.toJSON(), (data) => {
        fs.appendFile('results/Invoice' + new Date().valueOf() + '_' + t + '.xlsx', new Buffer(data), function (err) {
          if (err) {
            console.log(err);
          }else {
            if(t === times-1) {
              console.log('Export success');
              console.timeEnd(`test in ${times} times`);
            }
          }
        });
    }, (err) => {
        console.log(err);
    }, { useArrayBuffer: true });
}

runPerformance(1000)
```
完整的Demo工程请参考附件：spreadjs-nodejs-performance.zip
运行方式：
npm install
node ./app.js
运行结果：
![输入图片说明](https://gcdn.grapecity.com.cn/data/attachment/forum/202102/28/022349nm79kk6z6qfxz2fk.png)


本机配置：i7-9750H & 32G

### 总结：
1、从性能上分析：
SpreadJS in Node.js在“擅长”的批量I/O方面也输给了 GcExcel for Java，一方面由于GcExcel性能确实非常优异，它在Java平台上运用了很多优秀成熟的解决方案，做到了同类产品中最一流的性能表现，同时在对Excel和SpreadJS两方面的功能支持上也十分全面，是我们在服务器端处理Excel文档的首选方案；
2、从优化编码难度上分析：
另一方面，Node.js中采用了V8引擎，可以认为是目前针对JavaScript性能最好的引擎，但语言平台的硬伤，也确实很难弥补。这里所说的语言层面的瓶颈，并非是指JS不可能达到Java的性能表现，事实上如果你是V8引擎的核心开发者，完全有可能用JavaScript写出超过普通Java程序员的代码的性能，这里有一篇Vyacheslav Egorov大神用JS代码吊打Rust编译的WASM模块的文章。而是指要达到某个较高的性能指标，优化的难度和成本相比较而言，还是Java更低一些，从现实出发，毕竟不是每个程序员都能达到V8引擎核心开发者Vyacheslav Egorov的水平，而且我们还得分出大量的精力关注业务实现。
3、从技术选型上分析：
当然，本系列没有提到的一个对技术选型有决定性的地方就在于Node.js和Java/.Net平台还是有较大区别，如果项目本身采用的是Java Web或.Net Web架构，那选择GcExcel也是顺理成章的事情。另外认真阅读过第一篇的同学应该也注意到，Node.js中运行SpreadJS需要依赖诸如mock-browser / jsdom等第三方的组件，这对生产环境来说都是不可控的风险因素。
