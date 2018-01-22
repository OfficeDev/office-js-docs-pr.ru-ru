
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a>Создание надстройки Project, использующей REST с локальной службой OData Project Server

В этой статье описывается создание надстройки области задач для Project профессиональный 2013, которая сравнивает данные по материальным и трудовым затратам в активном проекте со средними значениями из всех проектов в текущем экземпляре Project Web App. Надстройка использует REST с библиотекой jQuery для получения доступа к службе отчетов OData **ProjectData** в Project Server 2013.


Код в данной статье основан на примере, разработанном Саурабхом Сангхви (Saurabh Sanghvi) и Эрвиндом Лаиром (Arvind Iyer), сотрудниками корпорации Майкрософт.

## <a name="prerequisites-for-creating-a-task-pane-add-in-that-reads-project-server-reporting-data"></a>Необходимые условия для создания надстроек области задач, читающей данные отчетов Project Server


Далее приводятся необходимые условия для создания надстройки области задач Project, считывающей данные из службы **ProjectData** в экземпляре Project Web App локальной установки Project Server 2013:


- Проверьте, что на локальном компьютере разработчика установлены самые последние пакеты обновления и обновления Windows. Операционной системой может быть Windows 7, Windows 8, Windows Server 2008 или Windows Server 2012.
    
- Project профессиональный 2013 требуется для подключения к Project Web App. На компьютере разработчика должен быть установлен Project профессиональный 2013, чтобы включить отладку по клавише **F5** с помощью Visual Studio.
    
     >**Примечание**. С помощью Project стандартный 2013 можно размещать надстройки области задач, но невозможно войти в Project Web App.
- Visual Studio 2015 с Инструменты разработчика Office для Visual Studio содержит шаблоны, позволяющие создавать Надстройки Office и SharePoint. Убедитесь, что у вас установлена самая последняя версия Office Developer Tools. См. раздел _Средства_ статьи [Надстройки Office и скачиваемые файлы для SharePoint](http://msdn.microsoft.com/en-us/office/apps/fp123627.aspx).
    
- Процедуры и примеры кода, приведенные в этой статье, получают доступ к службе **ProjectData**, предоставляемой Project Server 2013 в локальном домене. Методы jQuery в этой статье не работают с Project Online.
    
    Убедитесь, что служба **ProjectData** доступна на компьютере разработчика.
    

### <a name="procedure-1-to-verify-that-the-projectdata-service-is-accessible"></a>Процедура 1. Проверка доступности службы ProjectData


Чтобы разрешить браузеру напрямую отображать XML-данные из запроса REST, отключите вид чтения канала. Дополнительные сведения о том, как это сделать в Internet Explorer, см. в процедуре 1, этап 4 в статье [Создание запросов веб-каналов OData для данных отчетов Project](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).
    
Отправьте запрос службе **ProjectData** с помощью браузера, используя следующий URL-адрес: **http://ServerName /ProjectServerName /_api/ProjectData**. Например, если `http://MyServer/pwa` — это экземпляр Project Web App, то в браузере будут показаны следующие результаты:
    
```xml
     <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/" 
        xmlns="http://www.w3.org/2007/app" 
        xmlns:atom="http://www.w3.org/2005/Atom">
        <workspace>
            <atom:title>Default</atom:title>
            <collection href="Projects">
                <atom:title>Projects</atom:title>
            </collection>
            <collection href="ProjectBaselines">
                <atom:title>ProjectBaselines</atom:title>
            </collection>
            <!-- ... and 33 more collection elements -->
        </workspace>
        </service>
```
3. Вам может потребоваться предоставить свои сетевые учетные данные, чтобы увидеть результаты. Если браузер показывает сообщение "Ошибка 403, доступ запрещен", то либо у вас либо нет разрешений на вход для заданного экземпляра Project Web App, либо имеется проблема сети, требующая помощи администратора.
    

## <a name="using-visual-studio-to-create-a-task-pane-add-in-for-project"></a>Создание надстройки области задач для Project с помощью Visual Studio

Инструменты разработчика Office для Visual Studio включает шаблон надстроек области задач для Project 2013. Если вы создаете решение с именем **HelloProjectOData**, оно содержит следующие два проекта Visual Studio:


- Проект надстройки получает имя решения. Оно включает в себя XML-файл манифеста для приложения и настраивается на целевую платформу .NET Framework 4.5. В процедуре 3 показаны шаги по изменению манифеста надстройки **HelloProjectOData**.
    
- Веб-проект получает имя **HelloProjectODataWeb**. Оно содержит файлы JavaScript веб-страниц, файлы CSS, рисунки, ссылки и файлы конфигурации для веб-контента в области задач. Веб-проект настраивается на конечную платформу .NET Framework 4. В процедуре 4 и процедуре 5 показано, как изменить эти файлы в веб-проекте, чтобы создать функциональность надстройки **HelloProjectOData**.
    

### <a name="procedure-2-to-create-the-helloprojectodata-add-in-for-project"></a>Процедура 2. Создание надстройки HelloProjectOData для Project


1. Запустите Visual Studio 2015 от имени администратора и выберите команду **Создать проект** на начальной странице.
    
2. В диалоговом окне **Новый проект** разверните узлы **Шаблоны** > **Visual C#** > **Office/SharePoint** и выберите **Надстройки Office**. Выберите **.NET Framework 4.5.2** в раскрывающемся списке в верхней части центральной панели, а затем выберите **Надстройка Office** (см. следующий снимок экрана).
    
3. Чтобы разместить оба проекта Visual Studio в одной папке, выберите **Создать каталог для решения** и найдите требуемое расположение.
    
4. В поле **Имя** введите HelloProjectOData и нажмите кнопку **ОК**.
    
    **Создание надстройки Office**

    ![Создание приложения для Office 2013](../images/pj15_HelloProjectOData_CreatingApp.png)

5. В диалоговом окне **Выбор типа надстройки** выберите пункт **Надстройка области задач** и нажмите кнопку **Далее** (см. следующий снимок экрана).
    
    **Выбор типа создаваемой надстройки**

    ![Выбор типа создаваемого приложения](../images/pj15_HelloProjectOData_ChooseProject.png)

6. В диалоговом окне **Выбор ведущих приложений** снимите все флажки, кроме флажка **Project** (см. следующий снимок экрана), а затем нажмите кнопку **Готово**.
    
    **Выбор ведущего приложения**

    ![Выбор Project в качестве единственного ведущего приложения](../images/b2144f2c-51f6-4e61-bc0d-972125c57031.png)
    
    С помощью Visual Studio можно создавать проекты **HelloProjectOdata** и **HelloProjectODataWeb**.
    
В папке **AddIn** (см. следующий снимок экрана) находится файл App.css для пользовательских стилей CSS. В дочерней папке **Home** находится файл Home.html, который содержит ссылки на CSS-файлы и файлы JavaScript, которые использует надстройка, и код HTML5 для надстройки. Кроме того, файл Home.js предназначен для пользовательского кода JavaScript. В папке **Scripts** находятся файлы библиотек jQuery. В дочерней папке **Office** находятся библиотеки JavaScript, например office.js и project-15.js, а также языковые библиотеки для стандартных строк в надстройках Office. В папке **Content** находится файл Office.css, который содержит стили по умолчанию для всех надстроек Office.

**Просмотр файлов веб-проекта по умолчанию в обозревателе решений**

![Просмотр файлов веб-проекта в обозревателе решений](../images/pj15_HelloProjectOData_InitialSolutionExplorer.png)

Манифест проекта **HelloProjectOData** — это файл HelloProjectOData.xml. Его можно изменить при необходимости, чтобы добавить описание надстройки, ссылку на значок, сведения о дополнительных языках и другие параметры. В процедуре 3 изменяется только отображаемое имя надстройки и описание и добавляется значок.

Дополнительные сведения о манифесте см. в статьях [XML-манифест надстроек для Office](../../docs/overview/add-in-manifests.md) и [Справка по схеме для манифестов надстроек Office (версия 1.1)](../overview/add-in-manifests.md).


### <a name="procedure-3-to-modify-the-add-in-manifest"></a>Процедура 3. Изменение манифеста надстройки


1. Откройте файл HelloProjectOData.xml в Visual Studio.
    
2. Отображаемое имя по умолчанию — это имя проекта Visual Studio ("HelloProjectOData"). Например, измените значение по умолчанию элемента **DisplayName** на значение"Hello ProjectData".
    
3. Описание по умолчанию — "HelloProjectOData". Например, измените значение по умолчанию элемента Description на "Test REST queries of the ProjectData service" (тестирование запросов REST службы ProjectData).
    
4. Добавьте значок для отображения в раскрывающемся списке **Надстройки Office** на вкладке **Проект** ленты. Вы можете добавить файл значка в решении Visual Studio или использовать URL-адрес значка. 

Ниже описано, как добавить файл значка в решение Visual Studio:
    
1. В **обозревателе решений** откройте папку Images.
    
2. Чтобы отображаться в раскрывающемся списке **Надстройки Office**, значок должен иметь размер 32 x 32 пикселя. Например, установите пакет SDK Project 2013, затем выберите папку **Images** и добавьте следующий файл из пакета SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`
    
    Можно использовать свое изображение 32 x 32 или скопировать изображение ![Значок для приложения HelloProjectOData](../images/pj15_HelloProjectData_NewIcon.jpg) в файл NewIcon.png, а затем добавить этот файл в папку `HelloProjectODataWeb\Images`.

3. В манифесте HelloProjectOData.xml добавьте элемент **IconUrl** под элементом **Description**. Значением URL-адреса значка является относительный путь на файл значка размером 32 x 32. Например, добавьте следующую строку: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**. Теперь файл манифеста HelloProjectOData.xml содержит следующий текст (ваше значение **Id** будет другим):

```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82 </Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />
    
        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
```

## <a name="creating-the-html-content-for-the-helloprojectodata-add-in"></a>Создание HTML-контента для надстройки HelloProjectOData

Надстройка **HelloProjectOData** — это пример, который содержит сообщения отладки и сообщения об ошибках. Она не предназначена для использования в рабочей среде. Перед началом написания кода HTML-контента разработайте пользовательский интерфейс и алгоритм работы пользователя с надстройкой, а также выделите функции JavaScript, взаимодействующие с HTML-кодом. Дополнительные сведения см. в статье[Рекомендации по проектированию надстроек Office](../../docs/design/add-in-design.md). 

В верхней части области задач размещается отображаемое имя надстройки, соответствующее значению элемента **DisplayName** в манифесте. Элемент **body** в файле HelloProjectOData.html содержит другие элементы пользовательского интерфейса:

- Подзаголовок, указывающий на общую функциональность или тип работы, например: **ODATA REST QUERY**.
    
- Кнопка **Get ProjectData Endpoint** вызывает функцию **setOdataUrl** для получения конечной точки службы **ProjectData** и отображения ее в текстовом поле. Если Project не подключен к Project Web App, надстройка вызовет обработчик ошибок для отображения всплывающего сообщения об ошибке.
    
- Кнопка **Compare All Projects** отключена до тех пор, пока надстройка не получит действительную конечную точку OData. Когда пользователь нажимает эту кнопку, она вызывает функцию **retrieveOData**, которая использует запрос REST для получения сведений о материальных и трудовых затратах проекта из службы **ProjectData**.
    
- Таблица отображает средние значения затрат проекта, фактических затрат, трудозатрат и процент выполнения. В таблице также сравниваются значения текущего активного проекта со средними. Если текущее значение больше среднего по всем проектам, значение отображается красным цветом. Если текущее значение меньше среднего, оно отображается зеленым цветом. Если текущее значение недоступно, в таблице отображается значение **NA** синим цветом.
    
    Функция **retrieveOData** вызывает функцию **parseODataResult**, которая вычисляет и отображает значения таблицы.
    
     >**Примечание.** В этом примере данные о затратах и работе для активного проекта являются производными опубликованных значений. При изменении значений файла Project в службу **ProjectData** не вносятся изменения до публикации проекта.


### <a name="procedure-4-to-create-the-html-content"></a>Процедура 4. Создание HTML-контента

1. В элементе **head** файла Home.html добавьте любые дополнительные элементы **link** для CSS-файлов, используемых в надстройке. Шаблон проекта Visual Studio содержит ссылку на файл App.css, который можно использовать для настраиваемых стилей CSS.
    
2. Добавьте любые дополнительные элементы **script** для библиотек JavaScript, используемых в надстройке. Шаблон проекта содержит ссылки на файлы jQuery- _[версия]_.js, office.js и MicrosoftAjax.js из папки **Scripts**.
    
     >**Примечание.** Перед развертыванием надстройки измените ссылку office.js и ссылку jQuery на ссылку сети доставки содержимого (CDN). Ссылка CDN предоставляет самую последнюю версию и обеспечивает оптимальную производительность.

    Надстройка **HelloProjectOData** также использует файл SurfaceErrors.js, с помощью которого во всплывающих сообщениях отображаются ошибки. Можно скопировать код из раздела _Надежное программирование_ статьи [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), а затем добавить файл SurfaceErrors.js в папку **Scripts\Office** проекта **HelloProjectODataWeb**.
    
    Ниже приведен обновленный HTML-код для элемента **head** с дополнительной строкой для файла SurfaceErrors.js.
    
```html
      <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8" />
      <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
      <title>Test ProjectData Service</title>
    
      <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    
      <!-- Add your CSS styles to the following file -->
      <link rel="stylesheet" type="text/css" href="../Content/App.css" />
    
      <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
      <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
      <script src="../Scripts/jquery-1.7.1.js"></script>
    
      <!-- Use the CDN reference to office.js when deploying your add-in. -->
      <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->
    
      <!-- Use the local script references for Office.js to enable offline debugging -->
      <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
      <script src="../Scripts/Office/1.0/Office.js"></script>
    
      <!-- Add your JavaScript to the following files -->
      <script src="../Scripts/HelloProjectOData.js"></script>
      <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
```

3. В элементе **body** удалите существующий код из шаблона и затем добавьте код для пользовательского интерфейса. Если элемент должен заполняться данными или изменяться оператором jQuery, элемент должен содержать уникальный атрибут **id**. В следующем коде атрибуты **id** для элементов **button**, **span** и **td** (определение ячейки таблицы), используемых функциями jQuery, показаны жирным шрифтом.
    
    С помощью приведенного ниже HTML-кода можно добавить графическое изображение (например, логотип компании). Можно использовать логотип на свой выбор или же скопировать файл NewLogo.png из скачанного пакета SDK для Project 2013, а затем с помощью **обозревателя решений** добавить файл в папку `HelloProjectODataWeb\Images`.
    


```HTML
      <body>
      <div id="SectionContent">
        <div id="odataQueries">
          ODATA REST QUERY
        </div>
        <div id="odataInfo">
          <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
          <br /><br />
          <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
          <br />
        </div>
        <div id="compareProjectData">
          <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">Compare All Projects</button>
          <br />
        </div>
      </div>
      <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
          <tr>
            <td class="heading_leftCol"></td>
            <td class="heading_midCol"><strong>Average</strong></td>
            <td class="heading_rightCol"><strong>Current</strong></td>
          </tr>
          <tr>
            <td class="row_leftCol"><strong>Project Cost</strong></td>
            <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
          </tr>
          <tr>
            <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
            <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
          </tr>
          <tr>
            <td class="row_leftCol"><strong>Project Work</strong></td>
            <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
          </tr>
          <tr>
            <td class="row_leftCol"><strong>Project % Complete</strong></td>
            <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
          </tr>
        </table>
      </div>
      <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
      <br />
      <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
```


## <a name="creating-the-javascript-code-for-the-add-in"></a>Создание кода JavaScript для надстройки


Шаблон надстройки области задач для Project содержит код инициализации по умолчанию, который предназначен для демонстрации базовых действий получения и записи данных в документе для типичных приложений Office 2013. Так как Project 2013 не поддерживает действия записи в активный проект, а надстройка **HelloProjectOData** не использует метод **getSelectedDataAsync**, то можно удалить скрипт в функции **Office.initialize** и удалить функцию **setData** и функцию **getData** в файле HelloProjectOData.js по умолчанию.

В JavaScript содержатся глобальные константы для запроса REST и глобальные переменные, используемые в нескольких функциях. Кнопка **Get ProjectData Endpoint (получить конечную точку ProjectData)** вызывает функцию **setOdataUrl**, инициализирующую глобальные переменные и определяющую, подключен ли Project к Project Web App.

Оставшаяся часть файла HelloProjectOData.js содержит две функции: parseODataResult и retrieveOData. Функция **retrieveOData** вызывается когда пользователь выбирает команду **Compare All Projects (сравнить все проекты)**. Функция **parseODataResult** вычисляет средние значения, а затем заполняет таблицу сравнения значениями, отформатированными в соответствии с цветом и единицами измерения.


### <a name="procedure-5-to-create-the-javascript-code"></a>Процедура 5. Создание кода JavaScript


1. Удалите весь код в файле HelloProjectOData.js по умолчанию и затем добавьте глобальные переменные и функцию **Office.initialize**. Имена переменных, написанные полностью заглавными буквами подразумевают, что они являются константами; они позже будут использоваться с переменной **_pwa** для создания запроса REST в этом примере.
    
```js
      var PROJDATA = "/_api/ProjectData";
    var PROJQUERY = "/Projects?";
    var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
    var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
    var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
    var _pwa;           // URL of Project Web App.
    var _projectUid;    // GUID of the active project.
    var _docUrl;        // Path of the project document.
    var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData
    
    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
        });
    }
```

2. Добавьте функцию **setOdataUrl** и связанные функции. Функция **setOdataUrl** вызывает **getProjectGuid** и **getDocumentUrl** для инициализации глобальных переменных. В [методе getProjectFieldAsync](http://dev.office.com/reference/add-ins/shared/projectdocument.getprojectfieldasync) анонимная функция для параметра _callback_ включает кнопку **Compare All Projects** (сравнить все проекты) с помощью метода **removeAttr** из библиотеки jQuery, а затем отображает URL-адрес службы **ProjectData**. Если Project не подключен к Project Web App, функция вызывает ошибку, которая отображает всплывающее сообщение об ошибке. Файл SurfaceErrors.js содержит метод **throwError**.
    
     >**Примечание.** Если вы работаете в Visual Studio на компьютере с Project Server, раскомментируйте код после строки, отвечающей за инициализацию глобальной переменной **_pwa**, чтобы можно было использовать клавишу **F5** для отладки. Чтобы использовать метод jQuery **ajax** во время отладки на компьютере с Project Server, следует задать **localhost** в качестве значения URL-адреса PWA. При работе в Visual Studio на удаленном компьютере URL-адрес **localhost** не требуется. Перед развертыванием надстройки закомментируйте этот код.

```js
      function setOdataUrl() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _pwa = String(asyncResult.value.fieldValue);
    
                    // If you debug with Visual Studio on a local Project Server computer, 
                    // uncomment the following lines to use the localhost URL.
                    //var localhost = location.host.split(":", 1);
                    //var pwaStartPosition = _pwa.lastIndexOf("/");
                    //var pwaLength = _pwa.length - pwaStartPosition;
                    //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                    //_pwa = location.protocol + "//" + localhost + pwaName;
    
                    if (_pwa.substring(0, 4) == "http") {
                        _odataUrl = _pwa + PROJDATA;
                        $("#compareProjects").removeAttr("disabled");
                        getProjectGuid();
                    }
                    else {
                        _odataUrl = "No connection!";
                        throwError(_odataUrl, "You are not connected to Project Web App.");
                    }
                    getDocumentUrl();
                    $("#projectDataEndPoint").text(_odataUrl);
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the GUID of the active project.
    function getProjectGuid() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _projectUid = asyncResult.value.fieldValue;
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    
    // Get the path of the project in Project web app, which is in the form <>\ProjectName .
    function getDocumentUrl() {
        _docUrl = "Document path:\r\n" + Office.context.document.url;
    }
```

3. Добавьте функцию **retrieveOData**, которая объединяет значения для запроса REST и затем вызывает функцию **ajax** в jQuery для получения запрошенных данных из службы **ProjectData**. Переменная **support.cors** позволяет производить межплатформенный обмен ресурсами (CORS) с функцией **ajax**. Если оператор **support.cors** пропущен или имеет значение **false**, функция **ajax** возвращает ошибку **No transport (нет передачи)**.
    
     >**Примечание**. Приведенный ниже код подходит для локального сервера Project Server 2013. В Project Online можно использовать OAuth для проверки подлинности на основе токенов. Дополнительные сведения см. в статье [Обход ограничений, связанных с принципом одинакового источника, в надстройках для Office](../../docs/develop/addressing-same-origin-policy-limitations.md).

    Для вызова **ajax** можно использовать параметр _headers_ или параметр _beforeSend_. Параметр _complete_ — анонимная функция, поэтому находится в той же области, что и переменные в **retrieveOData**. Функция для параметра _complete_ отображает результаты в элементе управления **odataText**, а также вызывает метод **parseODataResult** для анализа и отображения ответа JSON. Параметр _error_ указывает именованную функцию **getProjectDataErrorHandler**, которая пишет сообщение об ошибке для элемента управления **odataText**, а также использует метод **throwError**, чтобы отобразить всплывающее сообщение.
    


```js
      /****************************************************************
    * Functions to get and parse the Project Server reporting data.
    *****************************************************************/
    
    // Get data about all projects on Project Server, 
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();
    
        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project Online.
        $.support.cors = true;
    
        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json",
            data: "",      // Empty string for the optional data.
            //headers: { "Accept": accept },
            beforeSend: function (xhr) {
                xhr.setRequestHeader("ACCEPT", accept);
            },
            complete: function (xhr, textStatus) {
                // Create a message to display in the text box.
                var message = "\r\ntextStatus: " + textStatus +
                    "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                    "\r\nStatus: " + xhr.status +
                    "\r\nResponseText:\r\n" + xhr.responseText;
    
                // xhr.responseText is the result from an XmlHttpRequest, which 
                // contains the JSON response from the OData service.
                parseODataResult(xhr.responseText, _projectUid);
    
                // Write the document name, response header, status, and JSON to the odataText control.
                $("#odataText").text(_docUrl);
                $("#odataText").append("\r\nREST query:\r\n" + restUrl);
                $("#odataText").append(message);
    
                if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                    $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
                }
            },
            error: getProjectDataErrorHandler
        });
    }
    
    function getProjectDataErrorHandler(data, errorCode, errorMessage) {
        $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
          + errorMessage);
        throwError(errorCode, errorMessage);
    }
```

4. Добавьте метод **parseODataResult**, который десериализует и обрабатывает ответ JSON из службы OData. Метод **parseODataResult** вычисляет средние значения материальных и трудовых затрат с точностью до одного или двух десятичных знаков, форматирует значения необходимым цветом и добавляет единицу измерения (**$**, **hrs** или **%**), а затем отображает значения в заданных ячейках таблицы.
    
    Если GUID активного проекта соответствует значению **ProjectId**, переменной **myProjectIndex** присваивается индекс проекта. Если **myProjectIndex** указывает, что активный проект опубликован на сервере Project Server, метод **parseODataResult** форматирует и отображает данные о затратах и работе для этого проекта. Если активный проект не опубликован, значения для него отображаются как **НД** в синем цвете.
    


```js
      // Calculate the average values of actual cost, cost, work, and percent complete   
    // for all projects, and compare with the values for the current project.
    function parseODataResult(oDataResult, currentProjectGuid) {
        // Deserialize the JSON string into a JavaScript object.
        var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
        var len = res.d.results.length;
        var projActualCost = 0;
        var projCost = 0;
        var projWork = 0;
        var projPercentCompleted = 0;
        var myProjectIndex = -1;
        for (i = 0; i < len; i++) {
            // If the current project GUID matches the GUID from the OData query,  
            // store the project index.
            if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
                myProjectIndex = i;
            }
            projCost += Number(res.d.results[i].ProjectCost);
            projWork += Number(res.d.results[i].ProjectWork);
            projActualCost += Number(res.d.results[i].ProjectActualCost);
            projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
        }
        var avgProjCost = projCost / len;
        var avgProjWork = projWork / len;
        var avgProjActualCost = projActualCost / len;
        var avgProjPercentCompleted = projPercentCompleted / len;
        
        // Round off cost to two decimal places, and round off other values to one decimal place.
        avgProjCost = avgProjCost.toFixed(2);
        avgProjWork = avgProjWork.toFixed(1);
        avgProjActualCost = avgProjActualCost.toFixed(2);
        avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);
        
        // Display averages in the table, with the correct units. 
        document.getElementById("AverageProjectCost").innerHTML = "$"
            + avgProjCost;
        document.getElementById("AverageProjectActualCost").innerHTML
            = "$" + avgProjActualCost;
        document.getElementById("AverageProjectWork").innerHTML
            = avgProjWork + " hrs";
        document.getElementById("AverageProjectPercentComplete").innerHTML
            = avgProjPercentCompleted + "%";
            
        // Calculate and display values for the current project.
        if (myProjectIndex != -1) {
            var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
            var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
            var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
            var myProjPercentCompleted =
              Number(res.d.results[myProjectIndex].ProjectPercentCompleted);
              
            myProjCost = myProjCost.toFixed(2);
            myProjWork = myProjWork.toFixed(1);
            myProjActualCost = myProjActualCost.toFixed(2);
            myProjPercentCompleted = myProjPercentCompleted.toFixed(1);
            
            document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;
            
            if (Number(myProjCost) <= Number(avgProjCost)) {
                document.getElementById("CurrentProjectCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectCost").style.color = "red"
            }
            
            document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;
            
            if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
                document.getElementById("CurrentProjectActualCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectActualCost").style.color = "red"
            }
            
            document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";
            
            if (Number(myProjWork) <= Number(avgProjWork)) {
                document.getElementById("CurrentProjectWork").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectWork").style.color = "green"
            }
            
            document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";
            
            if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
                document.getElementById("CurrentProjectPercentComplete").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectPercentComplete").style.color = "green"
            }
        }
        else {
            document.getElementById("CurrentProjectCost").innerHTML = "NA";
            document.getElementById("CurrentProjectCost").style.color = "blue"
            
            document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
            document.getElementById("CurrentProjectActualCost").style.color = "blue"
            
            document.getElementById("CurrentProjectWork").innerHTML = "NA";
            document.getElementById("CurrentProjectWork").style.color = "blue"
            
            document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
            document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
        }
    }
```


## <a name="testing-the-helloprojectodata-add-in"></a>Тестирование надстройки HelloProjectOData


Для тестирования и отладки надстройки **HelloProjectOData** с помощью Visual Studio 2015 на компьютере разработки должен быть установлен Project профессиональный 2013. Для работы с различными тестовыми сценариями убедитесь, что можно выбрать открытие файлов Project на локальном компьютере или подключение к Project Web App. Например, выполните следующие действия.


1. Во вкладке **ФАЙЛ** на ленте выберите вкладку **Сведения** в представлении Backstage, а затем выберите **Управление учетными записями**.
    
2. В диалоговом окне **Учетные записи Project Web App** список **Доступные учетные записи** может содержать несколько учетных записей Project Web App помимо локальной учетной записи **Компьютер**. В разделе **Во время запуска** выберите **Выбрать учетную запись**.
    
3. Закройте Project, чтобы среда Visual Studio могла запустить его для отладки надстройки.
    
Базовые тесты должны быть следующие:


- Запустите приложение в Visual Studio и откройте опубликованный проект из Project Web App, содержащего данные о материальных и трудовых затратах. Убедитесь, что надстройка отображает конечную точку **ProjectData** и правильно отображает данные о материальных и трудовых затратах в таблице. Можно использовать выходные данные в элементе управления **odataText** для проверки запроса REST и других сведений.
    
- Запустите надстройку еще раз и выберите профиль локального компьютера с помощью диалогового окна **Вход** во время запуска Project. Откройте локальный MPP-файл и протестируйте надстройку. Убедитесь, что она отображает сообщение об ошибке при попытке получить конечную точку **ProjectData**.
    
- Запустите надстройку еще раз и создайте проект, содержащий задачи с данными о материальных и трудовых затратах. Этот проект можно сохранить в Project Web App, но не публиковать. Убедитесь, что надстройка отображает данные с Project Server, но показывает **NA** для текущего проекта.
    

### <a name="procedure-6-to-test-the-add-in"></a>Процедура 6. Тестирование надстройки


1. Запустите Project профессиональный 2013, подключитесь к Project Web App и создайте тестовый проект. Назначьте задачи локальным ресурсам или ресурсам предприятия, настройте различные значения процента выполнения для некоторых задач и затем опубликуйте проект. Закройте Project, что позволит Visual Studio запустить Project для отладки надстройки.
    
2. В Visual Studio нажмите клавишу **F5**. Войдите в Project Web App и затем откройте проект, созданный на предыдущем шаге. Проект можно открыть в режиме чтения или в режиме редактирования.
    
3. На вкладке **Проект** ленты в раскрывающемся списке **Надстройки Office** выберите **Hello ProjectData** (см. рис. 4). Кнопка **Compare All Projects** должна быть отключена.
    
    **Рис. 4. Запуск надстройки HelloProjectOData**

    ![Тестирование приложения HelloProjectOData](../images/pj15_HelloProjectData_TestTheApp.png)

4. В области задач **Hello ProjectData** нажмите кнопку **Get ProjectData Endpoint (получить конечную точку ProjectData)**. Строка **projectDataEndPoint** должна показывать URL-адрес службы **ProjectData** и кнопка **Compare All Projects (сравнить все проекты)** должна быть включена (см. рис. 5).
    
5. Нажмите кнопку **Compare All Projects**. Надстройка может приостановить работу на время получения данных из службы **ProjectData**, а затем она должна отобразить отформатированные средние и текущие значения в таблице.
    
    **Рис. 5. Просмотр результатов запроса REST**

    ![Просмотр результатов запроса REST](../images/pj15_HelloProjectData_RESTresults.gif)

6. Проверьте выходные данные в текстовом поле. Они должны показывать путь к документу, запрос REST, сведения о состоянии и результаты JSON от вызовов **ajax** и **parseODataResult**. Выходные данные помогают понять, создать и отладить код в методе **parseODataResult**, такой как `projCost += Number(res.d.results[i].ProjectCost);`.
    
    Ниже приведен пример выходных данных для трех проектов в экземпляре Project Web App с разрывами строки и пробелами, добавленными для ясности.
    


```
          Document path: <>\WinProj test1
    
        REST query:
        http://sphvm-37189/pwa/_api/ProjectData/Projects?$filter=ProjectName ne 'Timesheet Administrative Work Items'
            &amp;$select=ProjectId, ProjectName, ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost
        
        textStatus: success
        ContentType: application/json;odata=verbose;charset=utf-8
        Status: 200
        
        ResponseText:
        {"d":{"results":[
        {"__metadata":
            {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
            "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
            "type":"ReportingData.Project"},
            "ProjectId":"ce3d0d65-3904-e211-96cd-00155d157123",
            "ProjectActualCost":"0.000000",
            "ProjectCost":"0.000000",
            "ProjectName":"Task list created in PWA",
            "ProjectPercentCompleted":0,
            "ProjectWork":"16.000000"},
        {"__metadata":
            {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
            "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
            "type":"ReportingData.Project"},
            "ProjectId":"c31023fc-1404-e211-86b2-3c075433b7bd",
            "ProjectActualCost":"700.000000",
            "ProjectCost":"2400.000000",
            "ProjectName":"WinProj test 2",
            "ProjectPercentCompleted":29,
            "ProjectWork":"48.000000"},
        {"__metadata":
            {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
            "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
            "type":"ReportingData.Project"},
            "ProjectId":"dc81fbb2-b801-e211-9d2a-3c075433b7bd",
            "ProjectActualCost":"1900.000000",
            "ProjectCost":"5200.000000",
            "ProjectName":"WinProj test1",
            "ProjectPercentCompleted":37,
            "ProjectWork":"104.000000"}
        ]}}
```

7. Остановите отладку (нажмите **SHIFT+F5**) и затем еще раз нажмите клавишу **F5** для запуска нового экземпляра Project. В диалоговом окне **Вход** выберите локальный профиль **Компьютер**, а не Project Web App. Создайте или откройте локальный MPP-файл проекта, откройте область задач **Hello ProjectData** и нажмите кнопку **Get ProjectData Endpoint**. Надстройка должна показать ошибку **No connection!** (см. рис. 6), а кнопка **Compare All Projects** должна остаться отключенной.
    
    **Рис. 6. Использование надстройки без подключения Project Web App**

    ![Использование приложения без подключения Project Web App](../images/pj15_HelloProjectData_NoConnection.gif)

8. Остановите отладку и нажмите клавишу **F5** снова. Войдите в Project Web App и создайте проект, содержащий данные о материальных и трудовых затратах. Проект можно сохранить, но не публикуйте его.
    
    Когда вы выбираете **Сравнить все проекты** в области задач **Hello ProjectData**, в полях столбца **Текущее** должны появиться значения **НД** в синем цвете (см. рис. 7).
    

    **Рис. 7. Сравнение неопубликованного проекта с другими проектами**

    ![Сравнение неопубликованного проекта с другими проектами](../images/pj15_HelloProjectData_NotPublished.gif)

Даже если ваша надстройка работала правильно в предыдущих тестах, есть другие тесты, которые необходимо выполнить. Например:

- Откройте в Project Web App проект, который не содержит данных о материальных и трудовых затратах для задач. В полях столбца **Current (текущий)** должны отображаться нули.
    
- Протестируйте проект, не содержащий задачи.
    
- Если вы измените надстройку и опубликуете ее, необходимо запустить аналогичные тесты снова с опубликованной надстройкой. Другие вопросы см. в разделе [Дальнейшие действия](#next-steps).
    

 >**Примечание.** Имеются ограничения на объем данных, который может быть возвращен в одном запросе службы **ProjectData**; этот объем данных меняется для разных сущностей. Например, набор сущностей **Projects** имеет ограничение по умолчанию в 100 проектов на запрос, но набор сущностей **Risks** имеет ограничение по умолчанию 200. Для установки в рабочей среде код в примере **HelloProjectOData** необходимо изменить для поддержки запросов, содержащих более 100 проектов. Дополнительные сведения см. в разделе [Дальнейшие действия](#next-steps) и в разделе [Создание запросов веб-каналов OData для данных отчетов Project](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).


## <a name="example-code-for-the-helloprojectodata-add-in"></a>Пример кода для надстройки HelloProjectOData


 **Файл HelloProjectOData.html.** Следующий код содержится в файле `Pages\HelloProjectOData.html` проекта **HelloProjectODataWeb**:


```html
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8" />
      <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
      <title>Test ProjectData Service</title>
    
      <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    
      <!-- Add your CSS styles to the following file -->
      <link rel="stylesheet" type="text/css" href="../Content/App.css" />
    
      <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
      <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
      <script src="../Scripts/jquery-1.7.1.js"></script>
    
      <!-- Use the CDN reference to Office.js when deploying your add-in -->
      <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->
    
      <!-- Use the local script references for Office.js to enable offline debugging -->
      <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
      <script src="../Scripts/Office/1.0/Office.js"></script>
    
      <!-- Add your JavaScript to the following files -->
      <script src="../Scripts/HelloProjectOData.js"></script>
      <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
      <div id="SectionContent">
        <div id="odataQueries">
          ODATA REST QUERY
        </div>
        <div id="odataInfo">
          <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
          <br />
          <br />
          <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
          <br />
        </div>
        <div id="compareProjectData">
          <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">
            Compare All Projects</button>
          <br />
        </div>
      </div>
      <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
          <tr>
            <td class="heading_leftCol"></td>
            <td class="heading_midCol"><strong>Average</strong></td>
            <td class="heading_rightCol"><strong>Current</strong></td>
          </tr>
          <tr>
            <td class="row_leftCol"><strong>Project Cost</strong></td>
            <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
          </tr>
          <tr>
            <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
            <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
          </tr>
          <tr>
            <td class="row_leftCol"><strong>Project Work</strong></td>
            <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
          </tr>
          <tr>
            <td class="row_leftCol"><strong>Project % Complete</strong></td>
            <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
          </tr>
        </table>
      </div>
      <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
      <br />
      <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
    </html>
```

    **HelloProjectOData.js file** The following code is in the `Scripts\Office\HelloProjectOData.js` file of the **HelloProjectODataWeb** project:




```js
    /* File: HelloProjectOData.js
    * JavaScript functions for the HelloProjectOData example task pane app.
    * October 2, 2012
    */
    
    var PROJDATA = "/_api/ProjectData";
    var PROJQUERY = "/Projects?";
    var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
    var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
    var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
    var _pwa;           // URL of Project Web App.
    var _projectUid;    // GUID of the active project.
    var _docUrl;        // Path of the project document.
    var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData
    
    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
        });
    }
    
    // Set the global variables, enable the Compare All Projects button,
    // and display the URL of the ProjectData service.
    // Display an error if Project is not connected with Project Web App.
    function setOdataUrl() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _pwa = String(asyncResult.value.fieldValue);
    
                    // If you debug with Visual Studio on a local Project Server computer, 
                    // uncomment the following lines to use the localhost URL.
                    //var localhost = location.host.split(":", 1);
                    //var pwaStartPosition = _pwa.lastIndexOf("/");
                    //var pwaLength = _pwa.length - pwaStartPosition;
                    //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                    //_pwa = location.protocol + "//" + localhost + pwaName;
    
                    if (_pwa.substring(0, 4) == "http") {
                        _odataUrl = _pwa + PROJDATA;
                        $("#compareProjects").removeAttr("disabled");
                        getProjectGuid();
                    }
                    else {
                        _odataUrl = "No connection!";
                        throwError(_odataUrl, "You are not connected to Project Web App.");
                    }
                    getDocumentUrl();
                    $("#projectDataEndPoint").text(_odataUrl);
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    
    // Get the GUID of the active project.
    function getProjectGuid() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _projectUid = asyncResult.value.fieldValue;
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    
    // Get the path of the project in Project web app, which is in the form <>\ProjectName .
    function getDocumentUrl() {
        _docUrl = "Document path:\r\n" + Office.context.document.url;
    }
    
    /****************************************************************
    * Functions to get and parse the Project Server reporting data.
    *****************************************************************/
    
    // Get data about all projects on Project Server, 
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();
    
        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project Online.
        $.support.cors = true;
    
        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json",
            data: "",      // Empty string for the optional data.
            //headers: { "Accept": accept },
            beforeSend: function (xhr) {
                xhr.setRequestHeader("ACCEPT", accept);
            },
            complete: function (xhr, textStatus) {
                // Create a message to display in the text box.
                var message = "\r\ntextStatus: " + textStatus +
                    "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                    "\r\nStatus: " + xhr.status +
                    "\r\nResponseText:\r\n" + xhr.responseText;
    
                // xhr.responseText is the result from an XmlHttpRequest, which 
                // contains the JSON response from the OData service.
                parseODataResult(xhr.responseText, _projectUid);
    
                // Write the document name, response header, status, and JSON to the odataText control.
                $("#odataText").text(_docUrl);
                $("#odataText").append("\r\nREST query:\r\n" + restUrl);
                $("#odataText").append(message);
    
                if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                    $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
                }
            },
            error: getProjectDataErrorHandler
        });
    }
    
    function getProjectDataErrorHandler(data, errorCode, errorMessage) {
        $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
          + errorMessage);
        throwError(errorCode, errorMessage);
    }
    
    // Calculate the average values of actual cost, cost, work, and percent complete   
    // for all projects, and compare with the values for the current project.
    function parseODataResult(oDataResult, currentProjectGuid) {
        // Deserialize the JSON string into a JavaScript object.
        var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
        var len = res.d.results.length;
        var projActualCost = 0;
        var projCost = 0;
        var projWork = 0;
        var projPercentCompleted = 0;
        var myProjectIndex = -1;
    
        for (i = 0; i < len; i++) {
            // If the current project GUID matches the GUID from the OData query,  
            // then store the project index.
            if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
                myProjectIndex = i;
            }
            projCost += Number(res.d.results[i].ProjectCost);
            projWork += Number(res.d.results[i].ProjectWork);
            projActualCost += Number(res.d.results[i].ProjectActualCost);
            projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
    
        }
        var avgProjCost = projCost / len;
        var avgProjWork = projWork / len;
        var avgProjActualCost = projActualCost / len;
        var avgProjPercentCompleted = projPercentCompleted / len;
    
        // Round off cost to two decimal places, and round off other values to one decimal place.
        avgProjCost = avgProjCost.toFixed(2);
        avgProjWork = avgProjWork.toFixed(1);
        avgProjActualCost = avgProjActualCost.toFixed(2);
        avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);
    
        // Display averages in the table, with the correct units. 
        document.getElementById("AverageProjectCost").innerHTML = "$"
            + avgProjCost;
        document.getElementById("AverageProjectActualCost").innerHTML
            = "$" + avgProjActualCost;
        document.getElementById("AverageProjectWork").innerHTML
            = avgProjWork + " hrs";
        document.getElementById("AverageProjectPercentComplete").innerHTML
            = avgProjPercentCompleted + "%";
    
        // Calculate and display values for the current project.
        if (myProjectIndex != -1) {
    
            var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
            var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
            var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
            var myProjPercentCompleted = Number(res.d.results[myProjectIndex].ProjectPercentCompleted);
    
            myProjCost = myProjCost.toFixed(2);
            myProjWork = myProjWork.toFixed(1);
            myProjActualCost = myProjActualCost.toFixed(2);
            myProjPercentCompleted = myProjPercentCompleted.toFixed(1);
    
            document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;
    
            if (Number(myProjCost) <= Number(avgProjCost)) {
                document.getElementById("CurrentProjectCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectCost").style.color = "red"
            }
    
            document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;
    
            if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
                document.getElementById("CurrentProjectActualCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectActualCost").style.color = "red"
            }
    
            document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";
    
            if (Number(myProjWork) <= Number(avgProjWork)) {
                document.getElementById("CurrentProjectWork").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectWork").style.color = "green"
            }
    
            document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";
    
            if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
                document.getElementById("CurrentProjectPercentComplete").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectPercentComplete").style.color = "green"
            }
        }
        else {    // The current project is not published.
            document.getElementById("CurrentProjectCost").innerHTML = "NA";
            document.getElementById("CurrentProjectCost").style.color = "blue"
    
            document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
            document.getElementById("CurrentProjectActualCost").style.color = "blue"
    
            document.getElementById("CurrentProjectWork").innerHTML = "NA";
            document.getElementById("CurrentProjectWork").style.color = "blue"
    
            document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
            document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
        }
    }
```

 **Файл App.css.** Следующий код содержится в файле `Content\App.css` проекта **HelloProjectODataWeb**:




```
/*
*  File: App.css for the HelloProjectOData app.
*  Updated: 10/2/2012
*/
 
body
{
    font-size: 11pt;
}
h1 
{
    font-size: 22pt;
}
h2 
{
    font-size: 16pt;
}

/******************************************************************
Code label class
******************************************************************/

.rest 
{
    font-family: 'Courier New';
    font-size: 0.9em;
}

/******************************************************************
Button classes
******************************************************************/

.button-wide {
    width: 210px;
    margin-top: 2px;
}
.button-narrow 
{
    width: 80px;
    margin-top: 2px;
}

/******************************************************************
Table styles
******************************************************************/

.infoTable
{
    text-align: center; 
    vertical-align: middle
}
.heading_leftCol
{
    width: 20px;
    height: 20px;
}
.heading_midCol
{
    width: 100px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.heading_rightCol
{
    width: 101px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.row_leftCol
{
    width: 20px;
    font-size: small; 
    font-weight: bold; 
}
.row_midCol
{
    width: 100px;
}
.row_rightCol
{
    width: 101px;
}
.logo
{
    width: 135px;
    height: 53px;
}
```

 **Файл SurfaceErrors.js.** Код для файла SurfaceErrors.js можно скопировать из раздела _ Надежное программирование_ статьи [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="next-steps"></a>Дальнейшие действия


Если бы надстройка **HelloProjectOData** была рабочей надстройкой, предназначенной для продажи в Магазин Office или распространения в каталоге надстроек SharePoint, она конструировалась бы по-другому. Например, здесь не было бы отладочных выходных данных в текстовом поле и, вероятно, не было бы кнопки для получения конечной точки **ProjectData**. Вам также следовало бы переписать функцию **retireveOData** для поддержки экземпляров Project Web App, содержащих более 100 проектов.

Надстройка должна содержать дополнительные проверки ошибок, а также логику для записи, объяснения или демонстрации пограничных случаев. Например, если экземпляр Project Web App содержит 1000 проектов со средней продолжительностью в пять дней и средними затратами в $2400, а активный проект является единственным с продолжительностью более 20 дней, то сравнение материальных и трудовых затрат может быть перекошено. Это может быть показано с помощью частотной диаграммы. Вам необходимо добавить команды для отображения продолжительности, сравнения проектов с одинаковой продолжительностью или сравнения проектов из одного или разных отделов. Либо добавить возможность пользователю выбирать из списка полей, которые требуется отобразить.

Для других запросов службы **ProjectData** имеются ограничения на длину строки запроса, что влияет на число шагов, которые запрос может предпринять для выборки из родительской коллекции в объект в дочерней коллекции. Например, двухшаговый запрос **Projects** в **Tasks** для получения элементов задач работает, но трехшаговый запрос, такой как **Projects** в **Tasks** в **Assignments**, для получения элемента назначения может превысить максимальную длину URL-адреса по умолчанию. Дополнительные сведения см. в разделе [Создание запросов веб-каналов OData для данных отчетов Project](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).

Если вы изменяете надстройку **HelloProjectOData** для использования в рабочей среде, выполните следующие действия.


- В файле HelloProjectOData.html для лучшей производительности измените ссылку office.js из локального проекта на ссылку CDN:
    
```HTML
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
```

- Перепишите функцию **retrieveOData** для разрешения запросов, обрабатывающих более 100 проектов. Например, можно получить число проектов с помощью запроса `~/ProjectData/Projects()/$count` и использовать оператор _$skip_ и оператор _$top_ в запросе REST для получения данных проекта. Запустите несколько запросов в цикле и затем усредните данные из всех запросов. Каждый запрос данных проекта будет выглядеть следующим образом: `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`.
    
    For more information, see [OData System Query Options Using the REST Endpoint](http://msdn.microsoft.com/library/8a938b9b-7fdb-45a3-a04c-4d2d5cf2e353.aspx). You can also use the [Set-SPProjectOdataConfiguration](http://technet.microsoft.com/library/jj219516%28v=office.15%29.aspx) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](http://msdn.microsoft.com/library/1ed14ee9-1a1a-4960-9b66-c24ef92cdf6b%28Office.15%29.aspx).
    
- Сведения о развертывании надстройки см. в разделе [Публикация надстройки Office](../publish/publish.md).
    

## <a name="additional-resources"></a>Дополнительные ресурсы



- [Надстройки области задач для Project](../project/project-add-ins.md)
    
- [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
    
- [ProjectData — Справочник по службе Project OData](http://msdn.microsoft.com/library/1ed14ee9-1a1a-4960-9b66-c24ef92cdf6b%28Office.15%29.aspx)
    
- [XML-манифест надстроек для Office](../../docs/overview/add-in-manifests.md)
    
- [Публикация надстройки Office](../publish/publish.md)
    
