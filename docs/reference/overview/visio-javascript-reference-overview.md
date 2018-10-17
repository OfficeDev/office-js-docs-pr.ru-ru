# <a name="visio-javascript-api-overview"></a>Обзор API JavaScript для Visio

С помощью API JavaScript для Visio вы можете внедрять схемы Visio в SharePoint Online. Внедренный документ Visio — схема, которая хранится в библиотеке документов SharePoint и отображается на странице SharePoint. Чтобы внедрить документ Visio, отобразите его в элементе `<iframe>` HTML. После этого вы сможете программным способом работать с внедренным документом при помощи API JavaScript для Visio.

![Документ Visio в iframe на странице SharePoint вместе с веб-частью редактора сценариев](../images/visio-api-block-diagram.png)


API JavaScript для Visio позволяет следующее:

* взаимодействовать с элементами документа Visio как со страницами, так и фигурами;
* создавать визуальную разметку на холсте документа Visio;
* создавать специальные обработчики событий мыши для документа;
* предоставлять своему решению данные документа, такие как текст фигуры, данные фигуры и гиперссылки.

В этой статье описано, как использовать API JavaScript для Visio с Visio Online, чтобы создавать решения для SharePoint Online. В ней представлены ключевые элементы, понимание роли которых крайне важно при использовании API, такие как прокси-объекты JavaScript, **EmbeddedSession**, **RequestContext**, а также методы **sync()**, **Visio.run()** и **load()**. В приведенных ниже примерах кода показано применение этих элементов.

## <a name="embeddedsession"></a>EmbeddedSession

Объект EmbeddedSession инициализирует взаимодействие между фреймом разработчика и фреймом Visio Online.

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a>Visio.run(session, function(context) { batch })

Метод **Visio.run()** выполняет пакетный сценарий, совершающий действия с объектной моделью Visio. Пакетные команды включают определения локальных прокси-объектов JavaScript и методов **sync()**, синхронизирующих состояние объектов Visio и локальных объектов, а также разрешение обещания. Преимущество пакетной обработки запросов в методе **Visio.run()** состоит в том, что при разрешении обещания все отслеживаемые объекты страницы, выделенные во время выполнения, автоматически освобождаются.

Метод run принимает объект session и RequestContext и возвращает обещание (обычно это результат **context.sync()**). Пакетную операцию можно выполнить, не указывая ее в методе **Visio.run()**. Однако в этом случае все ссылки на объекты страницы требуют отслеживания и управления вручную.

## <a name="requestcontext"></a>RequestContext

Объект RequestContext облегчает запросы на приложение Visio. Поскольку фрейм разработчика и приложение Visio Online выполняются в двух разных iframe, объект RequestContext (контекст в следующем примере) требуется для доступа к Visio и связанным с ним объектам, таким как страницы и фигуры, из фрейма разработчика.

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a>Прокси-объекты

Объекты JavaScript для Visio, объявленные и использованные в надстройке, — это прокси-объекты для реальных объектов в документе Visio. Все действия над прокси-объектами не реализуются в Visio, а состояние документа Visio — в прокси-объектах, пока оно не будет синхронизировано. Состояние документа синхронизируется при выполнении `context.sync()`.

Например, локальный объект JavaScript getActivePage объявлен в качестве ссылки на выбранный диапазон. Это можно использовать для постановки в очередь настройки его свойств и вызова методов. Действия над такими объектами не реализуются до выполнения метода **sync()**.

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a>sync()

Метод **sync()** синхронизирует состояние прокси-объектов JavaScript и реальных объектов в Visio путем выполнения поставленных в очередь инструкций над контекстом и получения свойств загруженных объектов Office для их использования в коде. Этот метод возвращает обещание, которое выполняется после завершения синхронизации. 

## <a name="load"></a>load()

Метод **load()** используется для заполнения прокси-объектов, созданных на уровне JavaScript надстройки. При попытке получения объекта, такого как документ, сначала на уровне JavaScript создается локальный прокси-объект. Такой объект можно использовать для добавления в очередь настройки его свойств и вызова методов. Но для чтения свойств или связей объекта сначала необходимо вызвать методы **load()** и **sync()**. Метод load() использует свойства и связи, которые требуется загрузить при вызове метода **sync()**.

Ниже представлен синтаксис метода **load()**.

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. **properties** - это список имен свойств, которые должны быть загружены, заданные как строки с разделителями-запятыми или массив имен. Дополнительные сведения см. в описаниях методов **.load()** под каждым объектом.

2. **loadOption** указывает объект, описывающий свойства select, expand, top и skip. Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](/javascript/api/office/officeextension.loadoption).

## <a name="example-printing-all-shapes-text-in-active-page"></a>Пример: Печать текста всех фигур на активной странице

Приведенный ниже пример показывает, как распечатать значение текста фигуры из объекта фигур массива.
Метод **Visio.run()** содержит пакет инструкций. В рамках этого пакета создается прокси-объект, который ссылается на фигуры в активном документе.

Все эти команды ставятся в очередь и выполняются при вызове метода **ctx.sync()**. Метод **sync()** возвращает обещание, с помощью которого его можно связать с другими операциями.

```js
Visio.run(session, function (context) {
    var page = context.document.getActivePage();
    var shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++) {
            var shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a>Сообщения об ошибках

Ошибки возвращаются с помощью объекта ошибки, состоящего из кода и сообщения. В таблице ниже перечислены возможные ошибки.

| error.code            | error.message |
|-----------------------|----------------------------------------------------------------|
| InvalidArgument       | Аргумент недопустим, отсутствует или имеет неправильный формат. |
| GeneralException      | При обработке запроса возникла внутренняя ошибка. |
| NotImplemented        | Запрашиваемая функция не реализована.  |
| UnsupportedOperation  | Выполняемая операция не поддерживается. |
| AccessDenied          | Вы не можете выполнить запрашиваемую операцию. |
| ItemNotFound          | Запрашиваемый ресурс не существует. |

## <a name="get-started"></a>Начало работы

Пример в этом разделе можно использовать для начала работы. В этом примере показано, как программно отобразить текст выбранной фигуры в схеме Visio. Чтобы приступить к работе, создайте классическую страницу в SharePoint Online или отредактируйте существующую страницу. Добавьте веб-часть редактора сценариев на странице и скопируйте и вставьте приведенный ниже код.

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        var page = context.document.getActivePage();
        var shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(var i=0; i<shapes.items.length;i++) {
                var shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

После этого все, что требуется — это URL-адрес схемы Visio, с которой вы хотите работать. Просто загрузите схему Visio в SharePoint Online и откройте ее в Visio Online. Оттуда откройте диалоговое окно Внедрить и используйте URL-адрес Внедрить в приведенном выше примере.

![Скопируйте URL-адрес файла Visio из диалога Внедрить](../images/Visio-embed-url.png)

Если вы используете Visio Online в Режиме правки, откройте диалоговое окно Внедрить, выбрав **Файл** > **Поделиться** > **Внедрить**. Если вы используете Visio Online в режиме просмотра, откройте диалоговое окно Внедрить, выбрав '... а затем **Внедрить**.

## <a name="open-api-specifications"></a>Открытые спецификации API

Мы публикуем новые API на странице [Открытые спецификации API](../openspec.md), чтобы вы могли делиться своим мнением о них. Узнайте, над какими функциями мы работаем, и поделитесь своим мнением о спецификациях.

## <a name="visio-javascript-api-reference"></a>Ссылка на API JavaScript для Visio

Для получения подробной информации об API JavaScript для Visio см. Справочную документацию [Visio JavaScript API](/javascript/api/visio).
