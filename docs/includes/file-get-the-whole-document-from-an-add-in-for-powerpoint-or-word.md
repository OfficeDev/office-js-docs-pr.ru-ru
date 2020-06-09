<span data-ttu-id="6b409-p101">Можно создать Надстройка Office для отправки или публикации одним щелчком документа Word 2013 или PowerPoint 2013 в удаленное расположение. В данной статье показано, как создать простую надстройку области задач для PowerPoint 2013, которая получает все представление в виде объекта данных и отправляет эти данные на веб-сервер через запрос HTTP.</span><span class="sxs-lookup"><span data-stu-id="6b409-p101">You can create an Office Add-in to provide one-click sending or publishing of a Word 2013 or PowerPoint 2013 document to a remote location. This article demonstrates how to build a simple task pane add-in for PowerPoint 2013 that gets all of the presentation as a data object and sends that data to a web server via an HTTP request.</span></span>

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a><span data-ttu-id="6b409-103">Необходимые условия создания надстройки для PowerPoint или Word</span><span class="sxs-lookup"><span data-stu-id="6b409-103">Prerequisites for creating an add-in for PowerPoint or Word</span></span>

<span data-ttu-id="6b409-p102">В этой статье предполагается, что вы создаете надстройку области задач для PowerPoint или Word с помощью текстового редактора. Чтобы создать такую надстройку, необходимо создать указанные ниже файлы.</span><span class="sxs-lookup"><span data-stu-id="6b409-p102">This article assumes that you are using a text editor to create the task pane add-in for PowerPoint or Word. To create the task pane add-in, you must create the following files:</span></span>

- <span data-ttu-id="6b409-106">В общей сетевой папке или на веб-сервере необходимо иметь следующие файлы:</span><span class="sxs-lookup"><span data-stu-id="6b409-106">On a shared network folder or on a web server, you need the following files:</span></span>

    - <span data-ttu-id="6b409-107">HTML-файл (GetDoc_App.html), содержащий пользовательский интерфейс, а также ссылки на файлы JavaScript (включая office.js и host-specific.js) и CSS-файлы.</span><span class="sxs-lookup"><span data-stu-id="6b409-107">An HTML file (GetDoc_App.html) that contains the user interface plus links to the JavaScript files (including office.js and host-specific .js files) and Cascading Style Sheet (CSS) files.</span></span>

    - <span data-ttu-id="6b409-108">Файл JavaScript (GetDoc_App.js), содержащий алгоритм надстройки.</span><span class="sxs-lookup"><span data-stu-id="6b409-108">A JavaScript file (GetDoc_App.js) to contain the programming logic of the add-in.</span></span>

    - <span data-ttu-id="6b409-109">Файл CSS (Program.css) для размещения стилей и форматирования для надстройки.</span><span class="sxs-lookup"><span data-stu-id="6b409-109">A CSS file (Program.css) to contain the styles and formatting for the add-in.</span></span>

- <span data-ttu-id="6b409-p103">Файл XML-манифеста (GetDoc_App.xml) для надстройки, доступный в общей сетевой папке или каталоге надстроек. Файл манифеста должен указывать на расположение HTML-файла, упомянутого ранее.</span><span class="sxs-lookup"><span data-stu-id="6b409-p103">An XML manifest file (GetDoc_App.xml) for the add-in, available on a shared network folder or add-in catalog. The manifest file must point to the location of the HTML file mentioned previously.</span></span>

<span data-ttu-id="6b409-112">Вы также можете создать надстройку для PowerPoint с помощью [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) или [генератора Yeoman для надстроек Office](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) или для Word с помощью [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) или [генератора Yeoman для надстроек Office](../quickstarts/word-quickstart.md?tabs=yeomangenerator).</span><span class="sxs-lookup"><span data-stu-id="6b409-112">You can also create an add-in for PowerPoint by using [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) or the [Yeoman generator for Office Add-ins](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) or for Word by using [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) or [Yeoman generator for Office Add-ins](../quickstarts/word-quickstart.md?tabs=yeomangenerator).</span></span>

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a><span data-ttu-id="6b409-113">Основные понятия, позволяющие создавать надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="6b409-113">Core concepts to know for creating a task pane add-in</span></span>

<span data-ttu-id="6b409-p104">Прежде чем приступать к разработке этой надстройки для PowerPoint или Word, ознакомьтесь с созданием Надстройки Office и работой с HTTP-запросами. В этой статье не рассмотрен способ расшифровки текста из HTTP-запросов на веб-сервере, зашифрованного с помощью Base64.</span><span class="sxs-lookup"><span data-stu-id="6b409-p104">Before you begin creating this add-in for PowerPoint or Word, you should be familiar with building Office Add-ins and working with HTTP requests. This article does not discuss how to decode Base64-encoded text from an HTTP request on a web server.</span></span> 

## <a name="create-the-manifest-for-the-add-in"></a><span data-ttu-id="6b409-116">Создание манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="6b409-116">Create the manifest for the add-in</span></span>

<span data-ttu-id="6b409-117">Файл XML-манифеста надстройки для PowerPoint предоставляет важные сведения о надстройке: о том, в каких приложениях она может размещаться, расположение HTML-файла, имя и описание надстройки, а также многие другие характеристики.</span><span class="sxs-lookup"><span data-stu-id="6b409-117">The XML manifest file for the add-in for PowerPoint provides important information about the add-in: what applications can host it, the location of the HTML file, the add-in title and description, and many other characteristics.</span></span>

1. <span data-ttu-id="6b409-118">В текстовом редакторе добавьте следующий код в файл манифеста.</span><span class="sxs-lookup"><span data-stu-id="6b409-118">In a text editor, add the following code to the manifest file.</span></span>

    ```xml  
    <?xml version="1.0" encoding="utf-8" ?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:type="TaskPaneApp">
        <Id>[Replace_With_Your_GUID]</Id>
        <Version>1.0</Version>
        <ProviderName>[Provider Name]</ProviderName>
        <DefaultLocale>EN-US</DefaultLocale>
        <DisplayName DefaultValue="Get Doc add-in" />
        <Description DefaultValue="My get PowerPoint or Word document add-in." />
        <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
        <Host Name="Document" />
        <Host Name="Presentation" />
        </Hosts>
        <DefaultSettings>
        <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

2. <span data-ttu-id="6b409-119">Сохраните файл как GetDoc_App.xml в сетевую папку или каталог надстроек, используя кодировку UTF-8.</span><span class="sxs-lookup"><span data-stu-id="6b409-119">Save the file as GetDoc_App.xml using UTF-8 encoding to a network location or to an add-in catalog.</span></span>

## <a name="create-the-user-interface-for-the-add-in"></a><span data-ttu-id="6b409-120">Создание пользовательского интерфейса надстройки</span><span class="sxs-lookup"><span data-stu-id="6b409-120">Create the user interface for the add-in</span></span>

<span data-ttu-id="6b409-p105">Для пользовательского интерфейса надстройки вы можете использовать формат HTML-код, внесенный прямо в файл GetDoc_App.html. Программная логика и функции надстройки должны содержаться в файле JavaScript (например, GetDoc_App.js).</span><span class="sxs-lookup"><span data-stu-id="6b409-p105">For the user interface of the add-in, you can use HTML, written directly into the GetDoc_App.html file. The programming logic and functionality of the add-in must be contained in a JavaScript file (for example, GetDoc_App.js).</span></span>

<span data-ttu-id="6b409-123">Используйте следующую процедуру для создания простого пользовательского интерфейса надстройки, содержащего заголовок и одну кнопку.</span><span class="sxs-lookup"><span data-stu-id="6b409-123">Use the following procedure to create a simple user interface for the add-in that includes a heading and a single button.</span></span>

1. <span data-ttu-id="6b409-124">В новый файл, используя текстовый редактор, добавьте следующий HTML-код.</span><span class="sxs-lookup"><span data-stu-id="6b409-124">In a new file in the text editor, add the following HTML.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <title>Publish presentation</title>
            <link rel="stylesheet" type="text/css" href="Program.css" />
            <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js" type="text/javascript"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="GetDoc_App.js"></script>
        </head>
        <body>
        <form>
            <h1>Publish presentation</h1>
            <br />
            <div><input id='submit' type="button" value="Submit" /></div>
            <br />
            <div><h2>Status</h2> 
                <div id="status"></div>
            </div>
        </form>
        </body>
    </html>
    ```

2. <span data-ttu-id="6b409-125">Сохраните файл под именем GetDoc_App.html в сетевую папку или на веб-сервер, используя кодировку UTF-8.</span><span class="sxs-lookup"><span data-stu-id="6b409-125">Save the file as GetDoc_App.html using UTF-8 encoding to a network location or to a web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6b409-126">Убедитесь, что теги **head** надстройки содержат тег **script** с рабочей ссылкой на файл office.js.</span><span class="sxs-lookup"><span data-stu-id="6b409-126">Be sure that the **head** tags of the add-in contains a **script** tag with a valid link to the office.js file.</span></span> 

    <span data-ttu-id="6b409-p106">Мы будем использовать немного CSS, чтобы придать надстройке простой, но современный и профессиональный вид. Используйте CSS для определения стиля надстройки.</span><span class="sxs-lookup"><span data-stu-id="6b409-p106">We'll use some CSS to give the add-in a simple, yet modern and professional appearance. Use the following CSS to define the style of the add-in.</span></span>

3. <span data-ttu-id="6b409-129">В новый файл, используя текстовый редактор, добавьте следующий CSS-код.</span><span class="sxs-lookup"><span data-stu-id="6b409-129">In a new file in the text editor, add the following CSS.</span></span>

    ```css  
    body
    {
        font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
    }
    h1,h2
    {
        text-decoration-color:#4ec724;
    }
    input [type="submit"], input[type="button"]
    {
        height:24px;
        padding-left:1em;
        padding-right:1em;
        background-color:white;
        border:1px solid grey;
        border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0;
        cursor:pointer;
    }
    ```

4. <span data-ttu-id="6b409-130">Сохраните файл как Program.css в сетевую папку или на веб-сервер, где размещен файл GetDoc_App.html, используя кодировку UTF-8.</span><span class="sxs-lookup"><span data-stu-id="6b409-130">Save the file as Program.css using UTF-8 encoding to the network location or to the web server where the GetDoc_App.html file is located.</span></span>

## <a name="add-the-javascript-to-get-the-document"></a><span data-ttu-id="6b409-131">Добавление JavaScript для получения документа</span><span class="sxs-lookup"><span data-stu-id="6b409-131">Add the JavaScript to get the document</span></span>

<span data-ttu-id="6b409-132">В коде надстройки обработчик события [Office.initialize](/javascript/api/office) добавляет обработчик события нажатия кнопки **Submit** (Отправить), расположенной на форме, и информирует пользователя о том, что надстройка готова.</span><span class="sxs-lookup"><span data-stu-id="6b409-132">In the code for the add-in, a handler to the [Office.initialize](/javascript/api/office) event adds a handler to the click event of the **Submit** button on the form and informs the user that the add-in is ready.</span></span>

<span data-ttu-id="6b409-133">В следующем примере кода показан обработчик события для `Office.initialize` события вместе со вспомогательной функцией `updateStatus` для записи в состояние div.</span><span class="sxs-lookup"><span data-stu-id="6b409-133">The following code example shows the event handler for the `Office.initialize` event along with a helper function, `updateStatus`, for writing to the status div.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked
        $('#submit').click(function () {
            sendFile();
        });

        // Update status
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo.innerHTML += message + "<br/>";
}
```

<span data-ttu-id="6b409-134">Когда вы наберете кнопку " **послать** " в пользовательском интерфейсе, надстройка вызывает `sendFile` функцию, которая содержит вызов метода [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="6b409-134">When you choose the **Submit** button in the UI, the add-in calls the `sendFile` function, which contains a call to the [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) method.</span></span> <span data-ttu-id="6b409-135">`getFileAsync`Метод использует асинхронный шаблон, аналогичный другим методам в API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="6b409-135">The `getFileAsync` method uses the asynchronous pattern, similar to other methods in the JavaScript API for Office.</span></span> <span data-ttu-id="6b409-136">В нем есть один обязательный параметр _fileType_ и два необязательных параметра _options_ и _callback_.</span><span class="sxs-lookup"><span data-stu-id="6b409-136">It has one required parameter, _fileType_, and two optional parameters,  _options_ and _callback_.</span></span> 


<span data-ttu-id="6b409-137">Параметр _filetype_ ожидает один из трех констант в перечислении [filetype](/javascript/api/office/office.filetype) : `Office.FileType.Compressed` ("сжатый"), **Office. filetype. PDF** ("PDF") или **Office. filetype. Text** ("Text").</span><span class="sxs-lookup"><span data-stu-id="6b409-137">The  _fileType_ parameter expects one of three constants from the [FileType](/javascript/api/office/office.filetype) enumeration: `Office.FileType.Compressed` ("compressed"), **Office.FileType.PDF** ("pdf"), or **Office.FileType.Text** ("text").</span></span> <span data-ttu-id="6b409-138">PowerPoint supports only **Compressed** as an argument; Word supports all three.</span><span class="sxs-lookup"><span data-stu-id="6b409-138">PowerPoint supports only **Compressed** as an argument; Word supports all three.</span></span> <span data-ttu-id="6b409-139">При передаче **сжатия** для параметра _filetype_ `getFileAsync` метод возвращает документ как файл презентации PowerPoint 2013 (*pptx) или файл документа Word 2013 (* docx), создавая временную копию файла на локальном компьютере.</span><span class="sxs-lookup"><span data-stu-id="6b409-139">When you pass in **Compressed** for the _fileType_ parameter, the `getFileAsync` method returns the document as a PowerPoint 2013 presentation file (*.pptx) or Word 2013 document file (*.docx) by creating a temporary copy of the file on the local computer.</span></span>

<span data-ttu-id="6b409-140">`getFileAsync`Метод возвращает ссылку на файл в виде объекта [File](/javascript/api/office/office.file) .</span><span class="sxs-lookup"><span data-stu-id="6b409-140">The `getFileAsync` method returns a reference to the file as a [File](/javascript/api/office/office.file) object.</span></span> <span data-ttu-id="6b409-141">`File`Объект предоставляет четыре члена: свойство [size](/javascript/api/office/office.file#size) , свойство [sliceCount](/javascript/api/office/office.file#slicecount) , метод [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) и метод [closeAsync](/javascript/api/office/office.file#closeasync-callback-) .</span><span class="sxs-lookup"><span data-stu-id="6b409-141">The `File` object exposes four members: the [size](/javascript/api/office/office.file#size) property, [sliceCount](/javascript/api/office/office.file#slicecount) property, [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) method, and [closeAsync](/javascript/api/office/office.file#closeasync-callback-) method.</span></span> <span data-ttu-id="6b409-142">`size`Свойство возвращает число байтов в файле.</span><span class="sxs-lookup"><span data-stu-id="6b409-142">The `size` property returns the number of bytes in the file.</span></span> <span data-ttu-id="6b409-143">`sliceCount`Возвращает количество объектов [slice](/javascript/api/office/office.slice) (обсуждаемых далее в этой статье) в файле.</span><span class="sxs-lookup"><span data-stu-id="6b409-143">The `sliceCount` returns the number of [Slice](/javascript/api/office/office.slice) objects (discussed later in this article) in the file.</span></span>

<span data-ttu-id="6b409-144">Используйте приведенный ниже код, чтобы получить документ PowerPoint или Word в качестве `File` объекта с помощью `Document.getFileAsync` метода, а затем вызывает локально определенную `getSlice` функцию.</span><span class="sxs-lookup"><span data-stu-id="6b409-144">Use the following code to get the PowerPoint or Word document as a `File` object using the `Document.getFileAsync` method and then makes a call to the locally defined `getSlice` function.</span></span> <span data-ttu-id="6b409-145">Обратите внимание, что `File` объект, переменная счетчика и общее число фрагментов в файле передаются вместе в вызове `getSlice` в анонимном объекте.</span><span class="sxs-lookup"><span data-stu-id="6b409-145">Note that the `File` object, a counter variable, and the total number of slices in the file are passed along in the call to `getSlice` in an anonymous object.</span></span>

```js
// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
        });
}
```

<span data-ttu-id="6b409-146">Локальная функция вызывает `getSlice` `File.getSliceAsync` метод, чтобы получить срез из `File` объекта.</span><span class="sxs-lookup"><span data-stu-id="6b409-146">The local function `getSlice` makes a call to the `File.getSliceAsync` method to retrieve a slice from the `File` object.</span></span> <span data-ttu-id="6b409-147">`getSliceAsync`Метод возвращает `Slice` объект из коллекции фрагментов.</span><span class="sxs-lookup"><span data-stu-id="6b409-147">The `getSliceAsync` method returns a `Slice` object from the collection of slices.</span></span> <span data-ttu-id="6b409-148">Метод имеет два обязательных параметра: _sliceIndex_ и _callback_.</span><span class="sxs-lookup"><span data-stu-id="6b409-148">It has two required parameters, _sliceIndex_ and _callback_.</span></span> <span data-ttu-id="6b409-149">Параметр _sliceIndex_ принимает целое число в качестве индексатора в коллекцию фрагментов.</span><span class="sxs-lookup"><span data-stu-id="6b409-149">The  _sliceIndex_ parameter takes an integer as an indexer into the collection of slices.</span></span> <span data-ttu-id="6b409-150">Как и другие функции в API JavaScript для Office, этот `getSliceAsync` метод также использует функцию обратного вызова в качестве параметра для обработки результатов вызова метода.</span><span class="sxs-lookup"><span data-stu-id="6b409-150">Like other functions in the JavaScript API for Office, the `getSliceAsync` method also takes a callback function as a parameter to handle the results from the method call.</span></span>
<span data-ttu-id="6b409-151">`getSlice`объектный объект вызывает метод **File. getSliceAsync** для получения среза из объекта **File** .</span><span class="sxs-lookup"><span data-stu-id="6b409-151">ion `getSlice` makes a call to the **File.getSliceAsync** method to retrieve a slice from the **File** object.</span></span> <span data-ttu-id="6b409-152">Метод **getSliceAsync** возвращает объект **Slice** из коллекции фрагментов.</span><span class="sxs-lookup"><span data-stu-id="6b409-152">The **getSliceAsync** method returns a **Slice** object from the collection of slices.</span></span> <span data-ttu-id="6b409-153">Метод имеет два обязательных параметра: _sliceIndex_ и _callback_.</span><span class="sxs-lookup"><span data-stu-id="6b409-153">It has two required parameters, _sliceIndex_ and _callback_.</span></span> <span data-ttu-id="6b409-154">Параметр _sliceIndex_ принимает целое число в качестве индексатора в коллекцию фрагментов.</span><span class="sxs-lookup"><span data-stu-id="6b409-154">The  _sliceIndex_ parameter takes an integer as an indexer into the collection of slices.</span></span> <span data-ttu-id="6b409-155">Как и другие функции в API JavaScript для Office, метод **getSliceAsync** также использует функцию обратного вызова в качестве параметра для обработки результатов вызова метода.</span><span class="sxs-lookup"><span data-stu-id="6b409-155">Like other functions in the Office JavaScript API, the **getSliceAsync** method also takes a callback function as a parameter to handle the results from the method call.</span></span>

<span data-ttu-id="6b409-156">`Slice`Объект предоставляет доступ к данным, которые содержатся в файле.</span><span class="sxs-lookup"><span data-stu-id="6b409-156">The `Slice` object gives you access to the data contained in the file.</span></span> <span data-ttu-id="6b409-157">Если иное не указано в параметре _Options_ `getFileAsync` метода, `Slice` размер объекта равен 4 МБ.</span><span class="sxs-lookup"><span data-stu-id="6b409-157">Unless otherwise specified in the _options_ parameter of the `getFileAsync` method, the `Slice` object is 4 MB in size.</span></span> <span data-ttu-id="6b409-158">`Slice`Объект предоставляет три свойства: [size](/javascript/api/office/office.slice#size), [Data](/javascript/api/office/office.slice#data)и [index](/javascript/api/office/office.slice#index).</span><span class="sxs-lookup"><span data-stu-id="6b409-158">The `Slice` object exposes three properties: [size](/javascript/api/office/office.slice#size), [data](/javascript/api/office/office.slice#data), and [index](/javascript/api/office/office.slice#index).</span></span> <span data-ttu-id="6b409-159">`size`Свойство получает размер среза в байтах.</span><span class="sxs-lookup"><span data-stu-id="6b409-159">The `size` property gets the size, in bytes, of the slice.</span></span> <span data-ttu-id="6b409-160">`index`Свойство получает целое число, представляющее положение фрагмента в коллекции фрагментов.</span><span class="sxs-lookup"><span data-stu-id="6b409-160">The `index` property gets an integer that represents the slice's position in the collection of slices.</span></span>

```js
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

<span data-ttu-id="6b409-161">`Slice.data`Свойство возвращает необработанные данные файла в виде массива байтов.</span><span class="sxs-lookup"><span data-stu-id="6b409-161">The `Slice.data` property returns the raw data of the file as a byte array.</span></span> <span data-ttu-id="6b409-162">Если данные имеют текстовый формат (то есть XML или обычного текста), фрагмент содержит необработанный текст.</span><span class="sxs-lookup"><span data-stu-id="6b409-162">If the data is in text format (that is, XML or plain text), the slice contains the raw text.</span></span> <span data-ttu-id="6b409-163">Если вы передаете значение **Office. filetype. сжатый** для параметра _filetype_ объекта `Document.getFileAsync` , фрагмент содержит двоичные данные файла в виде массива байтов.</span><span class="sxs-lookup"><span data-stu-id="6b409-163">If you pass in **Office.FileType.Compressed** for the _fileType_ parameter of `Document.getFileAsync`, the slice contains the binary data of the file as a byte array.</span></span> <span data-ttu-id="6b409-164">В случае файла PowerPoint или Word фрагменты содержат массивы байтов.</span><span class="sxs-lookup"><span data-stu-id="6b409-164">In the case of a PowerPoint or Word file, the slices contain byte arrays.</span></span>

<span data-ttu-id="6b409-p114">Чтобы преобразовать данные массива байтов в строку с кодировкой Base64, вам необходимо применить собственную функцию (или использовать доступную библиотеку). Сведения о кодировании Base64 с помощью JavaScript см. в статье [Кодирование и декодирование Base64](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).</span><span class="sxs-lookup"><span data-stu-id="6b409-p114">You must implement your own function (or use an available library) to convert byte array data to a Base64-encoded string. For information about Base64 encoding with JavaScript, see [Base64 encoding and decoding](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).</span></span>

<span data-ttu-id="6b409-167">После преобразования данных в формат Base64 вы можете передать их на веб-сервер несколькими способами, в том числе в виде основного текста HTTP-запроса POST.</span><span class="sxs-lookup"><span data-stu-id="6b409-167">Once you have converted the data to Base64, you can then transmit it to a web server in several ways -- including as the body of an HTTP POST request.</span></span>

<span data-ttu-id="6b409-168">Добавьте следующий код для отправки фрагмента веб-службе.</span><span class="sxs-lookup"><span data-stu-id="6b409-168">Add the following code to send a slice to a web service.</span></span>

> [!NOTE]
> <span data-ttu-id="6b409-169">Этот код отправляет файл PowerPoint или Word на веб-сервер в нескольких фрагментах.</span><span class="sxs-lookup"><span data-stu-id="6b409-169">This code sends a PowerPoint or Word file to the web server in multiple slices.</span></span> <span data-ttu-id="6b409-170">Веб-сервер или служба должны добавить каждый отдельный срез в один файл, а затем сохранить его в формате PPTX или DOCX, прежде чем можно будет выполнять какие-либо манипуляции.</span><span class="sxs-lookup"><span data-stu-id="6b409-170">The web server or service must append each individual slice into a single file, and then save it as a .pptx or .docx file, before you can perform any manipulations on it.</span></span>

```js
function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST
        // request to the web server.
        request.send(fileData);
    }
}
```

<span data-ttu-id="6b409-171">Как следует из названия, `File.closeAsync` метод закрывает подключение к документу и освобождает ресурсы.</span><span class="sxs-lookup"><span data-stu-id="6b409-171">As the name implies, the `File.closeAsync` method closes the connection to the document and frees up resources.</span></span> <span data-ttu-id="6b409-172">Хотя сборщик мусора Надстройки Office в песочнице собирает недействующие ссылки на файлы, рекомендуется явно закрывать файлы после того, как код завершил работу с ними.</span><span class="sxs-lookup"><span data-stu-id="6b409-172">Although the Office Add-ins sandbox garbage collects out-of-scope references to files, it is still a best practice to explicitly close files once your code is done with them.</span></span> <span data-ttu-id="6b409-173">`closeAsync`Метод имеет один параметр _callback_, который задает функцию, которая вызывается при завершении вызова.</span><span class="sxs-lookup"><span data-stu-id="6b409-173">The `closeAsync` method has a single parameter, _callback_, that specifies the function to call on the completion of the call.</span></span>

```js
function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```