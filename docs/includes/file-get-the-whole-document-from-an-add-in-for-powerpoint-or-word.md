Можно создать Надстройка Office для отправки или публикации одним щелчком документа Word 2013 или PowerPoint 2013 в удаленное расположение. В данной статье показано, как создать простую надстройку области задач для PowerPoint 2013, которая получает все представление в виде объекта данных и отправляет эти данные на веб-сервер через запрос HTTP.

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>Необходимые условия создания надстройки для PowerPoint или Word

В этой статье предполагается, что вы создаете надстройку области задач для PowerPoint или Word с помощью текстового редактора. Чтобы создать надстройку области задач, необходимо создать следующие файлы.

- В общей сетевой папке или на веб-сервере вам потребуются следующие файлы.

  - HTML-файл (GetDoc_App.html), содержащий пользовательский интерфейс, а также ссылки на файлы JavaScript (включая office.js и файлы .js приложения) и CSS-файлы.

  - Файл JavaScript (GetDoc_App.js), содержащий алгоритм надстройки.

  - Файл CSS (Program.css) для размещения стилей и форматирования для надстройки.

- Файл XML-манифеста (GetDoc_App.xml) для надстройки, доступный в общей сетевой папке или каталоге надстроек. Файл манифеста должен указывать на расположение HTML-файла, упомянутого ранее.

Вы также можете создать надстройку для PowerPoint с помощью [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) или [генератора Yeoman](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) для надстроек Office или Word с помощью [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) или [генератора Yeoman](../quickstarts/word-quickstart.md?tabs=yeomangenerator) для надстроек Office.

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>Основные понятия, позволяющие создавать надстройки области задач

Прежде чем приступать к разработке этой надстройки для PowerPoint или Word, ознакомьтесь с созданием Надстройки Office и работой с HTTP-запросами. В этой статье не рассматривается декодирование текста в кодировке Base64 из HTTP-запроса на веб-сервере.

## <a name="create-the-manifest-for-the-add-in"></a>Создание манифеста надстройки

Файл XML-манифеста надстройки для PowerPoint предоставляет важные сведения о надстройке: о том, в каких приложениях она может размещаться, расположение HTML-файла, имя и описание надстройки, а также многие другие характеристики.

1. В текстовом редакторе добавьте следующий код в файл манифеста.

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

2. Сохраните файл как GetDoc_App.xml в сетевую папку или каталог надстроек, используя кодировку UTF-8.

## <a name="create-the-user-interface-for-the-add-in"></a>Создание пользовательского интерфейса надстройки

Для пользовательского интерфейса надстройки вы можете использовать формат HTML-код, внесенный прямо в файл GetDoc_App.html. Программная логика и функции надстройки должны содержаться в файле JavaScript (например, GetDoc_App.js).

Используйте следующую процедуру для создания простого пользовательского интерфейса надстройки, содержащего заголовок и одну кнопку.

1. В новый файл, используя текстовый редактор, добавьте следующий HTML-код.

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

2. Сохраните файл под именем GetDoc_App.html в сетевую папку или на веб-сервер, используя кодировку UTF-8.

    > [!NOTE]
    > Убедитесь, что теги **head** надстройки содержат тег **script** с рабочей ссылкой на файл office.js.

    Мы будем использовать немного CSS, чтобы придать надстройке простой, но современный и профессиональный вид. Используйте CSS для определения стиля надстройки.

3. В новый файл, используя текстовый редактор, добавьте следующий CSS-код.

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

4. Сохраните файл как Program.css в сетевую папку или на веб-сервер, где размещен файл GetDoc_App.html, используя кодировку UTF-8.

## <a name="add-the-javascript-to-get-the-document"></a>Добавление JavaScript для получения документа

В коде надстройки обработчик события [Office.initialize](/javascript/api/office) добавляет обработчик события нажатия кнопки **Submit** (Отправить), расположенной на форме, и информирует пользователя о том, что надстройка готова.

В следующем примере кода показан `Office.initialize` обработчик события вместе с вспомогательной функцией для `updateStatus`записи в раздел состояния.

```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked.
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
    statusInfo[0].innerHTML += message + "<br/>";
}
```

При нажатии **кнопки "** `sendFile` Отправить" в пользовательском интерфейсе надстройка вызывает функцию, содержащую вызов метода [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) . Этот `getFileAsync` метод использует асинхронный шаблон, аналогичный другим методам в API JavaScript для Office. В нем есть один обязательный параметр _fileType_ и два необязательных параметра _options_ и _callback_.

Параметр  _fileType_ ожидает одну из трех констант из перечисления [FileType](/javascript/api/office/office.filetype) : `Office.FileType.Compressed` ("compressed"), **Office.FileType.PDF** ("pdf") или **Office.FileType.Text** ("text"). Текущая поддержка типов файлов для каждой платформы указана в примечаниях [Document.getFileType](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) . При передаче сжатого для параметра _fileType_ `getFileAsync` метод возвращает документ в виде файла презентации PowerPoint 2013 (*.pptx) или файла документа Word 2013 (*.docx), создав временную копию файла на локальном компьютере.

Метод `getFileAsync` возвращает ссылку на файл в виде объекта [File](/javascript/api/office/office.file) . Объект `File` предоставляет четыре члена: свойство [size](/javascript/api/office/office.file#office-office-file-size-member) , [свойство sliceCount](/javascript/api/office/office.file#office-office-file-slicecount-member) , [метод getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) и [метод closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)) . Свойство `size` возвращает количество байтов в файле. Возвращает `sliceCount` количество объектов [Slice](/javascript/api/office/office.slice) (рассмотренных далее в этой статье) в файле.

Используйте следующий код, чтобы получить документ PowerPoint или Word `File` `Document.getFileAsync` в качестве объекта с помощью метода, а затем выполните вызов локально определенной функции `getSlice` . Обратите внимание `File` , что объект, переменная `getSlice` счетчика и общее количество срезов в файле передаются в вызове анонимного объекта.

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

Локальная функция `getSlice` вызывает метод `File.getSliceAsync` для получения среза из `File` объекта. Метод `getSliceAsync` возвращает объект из `Slice` коллекции срезов. Метод имеет два обязательных параметра: _sliceIndex_ и _callback_. Параметр _sliceIndex_ принимает целое число в качестве индексатора в коллекцию фрагментов. Как и другие методы в API JavaScript для Office, `getSliceAsync` метод также принимает функцию обратного вызова в качестве параметра для обработки результатов вызова метода.
ion `getSlice` вызывает метод **File.getSliceAsync** для получения среза из **объекта File** . Метод **getSliceAsync** возвращает объект **Slice** из коллекции фрагментов. Метод имеет два обязательных параметра: _sliceIndex_ и _callback_. Параметр _sliceIndex_ принимает целое число в качестве индексатора в коллекцию фрагментов. Как и другие методы в API JavaScript для Office, метод **getSliceAsync** также принимает функцию обратного вызова в качестве параметра для обработки результатов вызова метода.

Объект `Slice` предоставляет доступ к данным, содержащимся в файле. Если в параметре  `getFileAsync` параметров метода не указано иное, `Slice` размер объекта — 4 МБ. Объект `Slice` предоставляет три свойства: [размер](/javascript/api/office/office.slice#office-office-slice-size-member), [данные](/javascript/api/office/office.slice#office-office-slice-data-member) и [индекс](/javascript/api/office/office.slice#office-office-slice-index-member). Свойство `size` получает размер среза в байтах. Свойство `index` получает целое число, представляющее позицию среза в коллекции срезов.

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

Свойство `Slice.data` возвращает необработанные данные файла в виде массива байтов. Если данные имеют текстовый формат (то есть XML или обычного текста), фрагмент содержит необработанный текст. Если передать **файл Office.FileType.Compressed** для параметра _fileType_`Document.getFileAsync`, срез будет содержать двоичные данные файла в виде массива байтов. В случае файла PowerPoint или Word фрагменты содержат массивы байтов.

Чтобы преобразовать данные массива байтов в строку с кодировкой Base64, вам необходимо применить собственную функцию (или использовать доступную библиотеку). Сведения о кодировании Base64 с помощью JavaScript см. в статье [Кодирование и декодирование Base64](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).

После преобразования данных в формат Base64 вы можете передать их на веб-сервер несколькими способами, в том числе в виде основного текста HTTP-запроса POST.

Добавьте следующий код для отправки фрагмента веб-службе.

> [!NOTE]
> Этот код отправляет файл PowerPoint или Word на веб-сервер в нескольких срезах. Веб-сервер или служба должны добавить каждый отдельный срез в один файл, а затем сохранить его в виде .pptx или .docx файла, прежде чем можно будет выполнять какие-либо манипуляции с ним.

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

Как следует из названия, метод `File.closeAsync` закрывает подключение к документу и освобождает ресурсы. Хотя сборщик мусора Надстройки Office в песочнице собирает недействующие ссылки на файлы, рекомендуется явно закрывать файлы после того, как код завершил работу с ними. Метод `closeAsync` имеет один _параметр обратного_ вызова, который указывает функцию, вызываемую при завершении вызова.

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