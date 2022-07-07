---
title: Создание надстройки Project, использующей REST с локальной службой OData Project Server
description: Узнайте, как создать надстройку области задач для Project профессиональный 2013, которая сравнивает затраты и рабочие данные в активном проекте со средними значениями для всех проектов в текущем экземпляре Project Web App.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 9d1c24d738bc919e927c50d17f0e25ead4eba44c
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659852"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a>Создание надстройки Project, использующей REST с локальной службой OData Project Server

В этой статье описывается создание надстройки области задач для Project профессиональный 2013, которая сравнивает данные по материальным и трудовым затратам в активном проекте со средними значениями из всех проектов в текущем экземпляре Project Web App. Надстройка использует REST с библиотекой jQuery для доступа к службе **отчетов OData ProjectData** в Project Server 2013.

Код в данной статье основан на примере, разработанном Саурабхом Сангхви (Saurabh Sanghvi) и Эрвиндом Лаиром (Arvind Iyer), сотрудниками корпорации Майкрософт.

## <a name="prerequisites"></a>Предварительные требования

Ниже приведены предварительные требования для создания надстройки области задач Project, которая считывает службу **ProjectData** экземпляра Project Web App в локальной установке Project Server 2013.

- Проверьте, что на локальном компьютере разработчика установлены самые последние пакеты обновления и обновления Windows. Операционной системой может быть Windows 7, Windows 8, Windows Server 2008 или Windows Server 2012.

- Project профессиональный 2013 требуется для подключения к Project Web App. На компьютере разработки должен быть установлен Project профессиональный 2013, чтобы включить **отладку F5** с помощью Visual Studio.

    > [!NOTE]
    > Project стандартный 2013 также может содержать надстройки области задач, но не может войти в Project Web App.

- Visual Studio 2015 с Инструменты разработчика Office для Visual Studio содержит шаблоны, позволяющие создавать Надстройки Office и SharePoint. Убедитесь, что у вас установлена самая последняя версия Office Developer Tools. См. раздел _Средства_ статьи [Надстройки Office и скачиваемые файлы для SharePoint](https://developer.microsoft.com/office/docs).

- Процедуры и примеры кода, приведенные в этой статье, имеют доступ к службе **ProjectData** Project Server 2013 в локальном домене. Методы jQuery в этой статье не работают с Project в Интернете.

    Убедитесь, **что служба ProjectData** доступна с компьютера разработки.

### <a name="procedure-1-verify-that-the-projectdata-service-is-accessible"></a>Процедура 1. Проверка доступности службы ProjectData

1. Чтобы разрешить браузеру напрямую отображать XML-данные из запроса REST, отключите вид чтения канала. Сведения о том, как это сделать в Internet Explorer, см. в процедуре 1, шаге 4 в веб-каналах [OData запроса для данных отчетов Project](/previous-versions/office/project-odata/jj163048(v=office.15)).

2. Запросите **службу ProjectData** с помощью браузера со следующим URL-адресом: **http://ServerName /ProjectServerName /_api/ProjectData**. Например, если Project Web App экземпляром`http://MyServer/pwa`, браузер отображает следующие результаты.

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/"
        xmlns="https://www.w3.org/2007/app"
        xmlns:atom="https://www.w3.org/2005/Atom">
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

## <a name="use-visual-studio-to-create-a-task-pane-add-in-for-project"></a>Создание надстройки области задач для Project с помощью Visual Studio

Инструменты разработчика Office для Visual Studio включает шаблон надстроек области задач для Project 2013. Если вы создаете решение **с именем HelloProjectOData**, оно содержит следующие два проекта Visual Studio:

- Проект надстройки получает имя решения. Оно включает в себя XML-файл манифеста для приложения и настраивается на целевую платформу .NET Framework 4.5. В процедуре 3 показаны шаги по изменению манифеста для **надстройки HelloProjectOData** .

- Веб-проект называется **HelloProjectODataWeb**. Оно содержит файлы JavaScript веб-страниц, файлы CSS, рисунки, ссылки и файлы конфигурации для веб-контента в области задач. Веб-проект настраивается на конечную платформу .NET Framework 4. В процедуре 4 и процедуре 5 показано, как изменить эти файлы в веб-проекте, чтобы создать функциональность надстройки **HelloProjectOData**.

### <a name="procedure-2-create-the-helloprojectodata-add-in-for-project"></a>Процедура 2. Создание надстройки HelloProjectOData для Project

1. Запустите Visual Studio 2015 от имени администратора, а затем выберите "Создать **проект** " на начальной странице.

2. В **диалоговом** окне "Новый проект" разверните узлы **"** Шаблоны", **"** C#" и "**Office/SharePoint**", а затем выберите надстройки **Office**. Выберите **платформа .NET Framework 4.5.2** в раскрывающемся списке целевой платформы в верхней части центральной панели, а затем выберите надстройку **Office** (см. следующий снимок экрана).

3. Чтобы разместить оба проекта Visual Studio в одной папке, выберите **Создать каталог для решения** и найдите требуемое расположение.

4. В поле **"Имя** " введитеHelloProjectOData и нажмите кнопку **"ОК"**.

    *Рис. 1. Создание надстройки Office*

    ![Создайте надстройку Office.](../images/pj15-hello-project-o-data-creating-app.png)

5. В диалоговом окне **Выбор типа надстройки** выберите пункт **Надстройка области задач** и нажмите кнопку **Далее** (см. следующий снимок экрана).

    *Рис. 2. Выбор типа создаваемой надстройки*

    ![Выберите тип создаемой надстройки.](../images/pj15-hello-project-o-data-choose-project.png)

6. В диалоговом окне **Выбор ведущих приложений** снимите все флажки, кроме флажка **Project** (см. следующий снимок экрана), а затем нажмите кнопку **Готово**.

    *Рис. 3. Выбор ведущего приложения*

    ![Выберите Project в качестве только ведущего приложения.](../images/create-office-add-in.png)

    Visual Studio создает проект **HelloProjectOdata** и **проект HelloProjectODataWeb** .

**Папка AddIn** (см. следующий снимок экрана) содержит файл App.css для пользовательских стилей CSS. Во вложенной папке **Home** находится файл Home.html, содержащий ссылки на CSS-файлы и файлы JavaScript, используемые надстройкой, а также содержимое HTML5 для этой надстройки. Также в ней располагается файл Home.js, предназначенный для настраиваемого кода JavaScript. Папка **Scripts** содержит файлы библиотеки jQuery. Во вложенной папке **Office** находятся библиотеки JavaScript, например office.js и project-15.js, а также языковые библиотеки для стандартных строк в надстройках Office. В папке **Content** находится файл Office.css, содержащий стили по умолчанию для всех надстроек Office.

*Рис. 4. Просмотр файлов веб-проекта по умолчанию в Обозреватель решений*

![Просмотрите файлы веб-проекта в Обозреватель решений.](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

Манифест проекта **HelloProjectOData** — это файл HelloProjectOData.xml. Его можно изменить при необходимости, чтобы добавить описание надстройки, ссылку на значок, сведения о дополнительных языках и другие параметры. В процедуре 3 изменяется только отображаемое имя надстройки и описание и добавляется значок.

Дополнительные сведения о манифесте см. в статьях [XML-манифест надстроек для Office](../develop/add-in-manifests.md) и [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md#see-also).

### <a name="procedure-3-modify-the-add-in-manifest"></a>Процедура 3. Изменение манифеста надстройки

1. Откройте файл HelloProjectOData.xml в Visual Studio.

2. Отображаемое имя по умолчанию — это имя проекта Visual Studio ("HelloProjectOData"). Например, измените значение элемента по **\<DisplayName\>** умолчанию на "Hello ProjectData".

3. Описание по умолчанию — "HelloProjectOData". Например, измените значение по умолчанию элемента Description на "Test REST queries of the ProjectData service" (тестирование запросов REST службы ProjectData).

4. Добавьте значок для отображения в раскрывающемся списке **Надстройки Office** на вкладке **PROJECT** ленты. Можно добавить файл значка в решении Visual Studio или использовать URL-адрес значка. 

Ниже показано, как добавить файл значка в решение Visual Studio.

1. В **Обозреватель решений** перейдите в папку с именем Images.

2. Чтобы отображаться в раскрывающемся списке **Надстройки Office**, значок должен иметь размер 32 x 32 пикселя. Например, установите пакет SDK Project 2013, затем выберите папку **Images** и добавьте следующий файл из пакета SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`

    Кроме того, используйте собственный значок 32 x 32; или скопируйте следующее изображение в файл с именем NewIcon.png и добавьте этот файл в папку  `HelloProjectODataWeb\Images` .

    ![Значок приложения HelloProjectOData.](../images/pj15-hello-project-data-new-icon.jpg)

3. В HelloProjectOData.xml **\<IconUrl\>** **\<Description\>** добавьте элемент под элементом, где значение URL-адреса значка является относительным путем к файлу значка 32x32. Например, добавьте следующую строку: `<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />`. Теперь HelloProjectOData.xml манифеста содержит следующее (ваше **\<Id\>** значение будет другим):

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82</Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="create-the-html-content-for-the-helloprojectodata-add-in"></a>Создание HTML-содержимого для надстройки HelloProjectOData

**Надстройка HelloProjectOData** — это пример, который включает отладку и вывод ошибок. Он не предназначен для использования в рабочей среду. Прежде чем приступить к написанию кода HTML-содержимого, разработать пользовательский интерфейс и пользовательский интерфейс для надстройки и укажите функции JavaScript, взаимодействующие с HTML-кодом. Дополнительные сведения см. [в рекомендациях по проектированию надстроек Office](../design/add-in-design.md).

В верхней части области задач отображается отображаемое имя надстройки, **\<DisplayName\>** которое является значением элемента в манифесте. Элемент **body** в файле HelloProjectOData.html содержит другие элементы пользовательского интерфейса:

- Подзаголовок, указывающий на общую функциональность или тип работы, например: **ODATA REST QUERY**.

- **Кнопка "Получить конечную точку ProjectData**`setOdataUrl`" вызывает функцию, чтобы получить конечную точку службы **ProjectData** и отобразить ее в текстовом поле. Если Project не подключен к Project Web App, надстройка вызовет обработчик ошибок для отображения всплывающего сообщения об ошибке.

- Кнопка **Compare All Projects** отключена до тех пор, пока надстройка не получит действительную конечную точку OData. При нажатии кнопки вызывается `retrieveOData` функция, которая использует запрос REST для получения данных о затратах проекта и рабочих данных из **службы ProjectData** .

- Таблица отображает средние значения затрат проекта, фактических затрат, трудозатрат и процент выполнения. В таблице также сравниваются значения текущего активного проекта со средними. Если текущее значение больше среднего по всем проектам, значение отображается красным цветом. Если текущее значение меньше среднего, оно отображается зеленым цветом. Если текущее значение недоступно, в таблице отображается значение **NA** синим цветом.

    Функция `retrieveOData` вызывает функцию `parseODataResult` , которая вычисляет и отображает значения для таблицы.

    > [!NOTE]
    > В данном примере данные о материальных и трудовых затратах по активному проекту извлекаются из опубликованных значений. Если изменить значения в Project, служба **ProjectData** не будет знать об изменениях до тех пор, пока проект не опубликован.

### <a name="procedure-4-create-the-html-content"></a>Процедура 4. Создание HTML-содержимого

1. В элементе **head** файла Home.html добавьте любые дополнительные элементы **link** для CSS-файлов, используемых в надстройке. Шаблон проекта Visual Studio содержит ссылку на файл App.css, который можно использовать для настраиваемых стилей CSS.

2. Добавьте дополнительные **элементы** скрипта для библиотек JavaScript, которые использует надстройка. Шаблон проекта содержит ссылки на файлы jQuery _[version]_.js, office.js и MicrosoftAjax.js в **папке "Скрипты** ".

    > [!NOTE]
    > Перед развертыванием надстройки измените ссылку office.js и ссылку jQuery на ссылку сети доставки содержимого (CDN). Ссылка CDN предоставляет самую последнюю версию и обеспечивает оптимальную производительность.

    Надстройка **HelloProjectOData** использует файл SurfaceErrors.js, который отображает ошибки и всплывающее сообщение. Код можно скопировать из раздела "Надежное программирование" статьи "Создание первой надстройки области задач для [Project 2013](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)" с помощью текстового редактора, а затем добавить файл SurfaceErrors.js в папку **Scripts\Office** проекта **HelloProjectODataWeb**.

    Ниже приведен обновленный HTML-код для **головного** элемента с дополнительной строкой для SurfaceErrors.js файла.

    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

    <!-- Add your CSS styles to the following file. -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>

    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>

    <!-- Add your JavaScript to the following files. -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

3. В **элементе body** удалите существующий код из шаблона и добавьте код для пользовательского интерфейса. Если элемент должен заполняться данными или изменяться оператором jQuery, элемент должен содержать уникальный атрибут **id**. В следующем коде атрибуты **идентификатора** для элементов **button**, **span** и **td** (определение ячеек таблицы), используемых функциями jQuery, отображаются полужирным шрифтом.

   Следующий HTML-код добавляет графическое изображение, которое может быть эмблемой компании. Вы можете использовать логотип по вашему выбору или скопировать файл NewLogo.png из скачанного пакета SDK project 2013,  `HelloProjectODataWeb\Images` а затем с помощью Обозреватель решений добавить файл в папку.

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

## <a name="create-the-javascript-code-for-the-add-in"></a>Создание кода JavaScript для надстройки

Шаблон надстройки области задач для Project содержит код инициализации по умолчанию, который предназначен для демонстрации базовых действий получения и записи данных в документе для типичных приложений Office 2013. Так как Project 2013 не поддерживает действия, которые записывают в активный проект, а надстройка **HelloProjectOData** `getSelectedDataAsync` не использует этот метод, `Office.initialize` можно удалить скрипт в функции и `setData` `getData` удалить функцию и функцию в файле HelloProjectOData.js по умолчанию.

В JavaScript содержатся глобальные константы для запроса REST и глобальные переменные, используемые в нескольких функциях. **Кнопка "Получить конечную точку ProjectData**`setOdataUrl`" вызывает функцию, которая инициализирует глобальные переменные и определяет, подключен ли Project к Project Web App.

Остальная часть файла HelloProjectOData.js включает две функции: `retrieveOData` функция вызывается, когда пользователь выбирает "Сравнить все проекты **",**`parseODataResult` а функция вычисляет средние значения, а затем заполняет таблицу сравнения значениями, отформатированными для цвета и единиц.

### <a name="procedure-5-create-the-javascript-code"></a>Процедура 5. Создание кода JavaScript

1. Удалите весь код в файле HelloProjectOData.js по умолчанию, а затем добавьте глобальные переменные и `**`функцию Office.initialize. Имена переменных, написанные полностью заглавными буквами подразумевают, что они являются константами; они позже будут использоваться с переменной **_pwa** для создания запроса REST в этом примере.

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

2. Добавление `setOdataUrl` и связанные функции. Функция `setOdataUrl` вызывает и `getProjectGuid` инициализирует `getDocumentUrl` глобальные переменные. В [методе getProjectFieldAsync](/javascript/api/office/office.document) `removeAttr` анонимная функция для  параметра обратного вызова включает кнопку "Сравнить все проекты" с помощью метода в библиотеке jQuery, а затем отображает URL-адрес службы **ProjectData**. Если Project не подключен к Project Web App, функция вызывает ошибку, которая отображает всплывающее сообщение об ошибке. Метод SurfaceErrors.js файле `throwError` .

   > [!NOTE]
   > Если вы работаете с Visual Studio на компьютере Project Server, то для того, чтобы использовать отладку по клавише **F5**, раскомментируйте код после строки, инициализирующей глобальную переменную **_pwa**. Чтобы включить использование метода jQuery `ajax` при отладке на компьютере Project Server, `localhost` необходимо задать значение ДЛЯ URL-адреса PWA. При запуске Visual Studio на удаленном компьютере  `localhost` URL-адрес не требуется. Перед развертыванием надстройки закомментируйте этот код.

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

3. Добавьте функцию `retrieveOData` , которая сцепляет значения для запроса REST `ajax` , а затем вызывает функцию в jQuery, чтобы получить запрашиваемые данные из **службы ProjectData** . **Переменная support.cors** обеспечивает совместное использование ресурсов независимо от источника (CORS) с помощью функции`ajax`. Если **инструкция support.cors** отсутствует или имеет значение **false**, `ajax` функция возвращает ошибку **"Нет транспорта** ".

   > [!NOTE]
   > Приведенный ниже код подходит для локального сервера Project Server 2013. В Project в Интернете можно использовать OAuth для проверки подлинности на основе токенов. Дополнительные сведения см. в статье [Обход ограничений, связанных с принципом одинакового источника, в надстройках Office](../develop/addressing-same-origin-policy-limitations.md).

   В вызове `ajax` можно использовать параметр _заголовков_ или _параметр beforeSend_ . Полный _параметр_ является анонимной функцией, поэтому он находится в той же области, что и переменные в `retrieveOData`. Функция для полного  _параметра отображает_ `odataText` `parseODataResult` результаты в элементе управления, а также вызывает метод для анализа и отображения ответа JSON. Параметр _ошибки_ указывает именованную `getProjectDataErrorHandler` функцию, `odataText` `throwError` которая записывает сообщение об ошибке в элемент управления, а также использует метод для отображения всплывающего сообщения.

    ```js
    // Functions to get and parse the Project Server reporting data./

    // Get data about all projects on Project Server,
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();

        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project on the web.
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

4. Добавьте метод `parseODataResult` , который десериализует и обрабатывает ответ JSON из службы OData. `parseODataResult` Метод вычисляет средние значения затрат и рабочих данных с точностью до одного или двух десятичных разрядов, форматирует значения с правильным цветом и добавляет единицу (**$**, **hrs****%** или), а затем отображает значения в указанных ячейках таблицы.

   Если GUID активного проекта соответствует `ProjectId` значению, `myProjectIndex` переменной задается индекс проекта. Если `myProjectIndex` указывает, что активный проект опубликован на Project Server, `parseODataResult` метод форматирует и отображает данные о затратах и трудоемких данных для этого проекта. Если активный проект не опубликован, то для него отображается значение **NA** синим цветом.

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

## <a name="test-the-helloprojectodata-add-in"></a>Тестирование надстройки HelloProjectOData

Чтобы протестировать и отладить надстройку **HelloProjectOData** в Visual Studio 2015, на компьютере разработки необходимо установить Project профессиональный 2013. Для работы с различными тестовыми сценариями убедитесь, что можно выбрать открытие файлов Project на локальном компьютере или подключение к Project Web App. Например, выполните следующие действия.

1. На **вкладке "** Файл" **выберите вкладку** "Сведения" в представлении Backstage, а затем выберите "Управление **учетными записями"**.

2. В **диалоговом** окне "Учетные записи веб-приложения Project" список "Доступные учетные записи" может содержать несколько учетных записей Project Web App учетных записей в дополнение к учетной записи **локального** компьютера. В разделе **Во время запуска** выберите **Выбрать учетную запись**.

3. Закройте Project, чтобы среда Visual Studio могла запустить его для отладки надстройки.

Базовые тесты должны быть следующие:

- Запустите приложение в Visual Studio и откройте опубликованный проект из Project Web App, содержащего данные о материальных и трудовых затратах. Убедитесь, что надстройка отображает конечную точку **ProjectData** и правильно отображает данные о затратах и работе в таблице. Можно использовать выходные данные в элементе управления **odataText** для проверки запроса REST и других сведений.

- Запустите надстройку еще раз и выберите профиль локального компьютера с помощью диалогового окна **Вход** во время запуска Project. Откройте локальный MPP-файл и протестируйте надстройку. Убедитесь, что она отображает сообщение об ошибке при попытке получить конечную точку **ProjectData**.

- Запустите надстройку еще раз и создайте проект, содержащий задачи с данными о материальных и трудовых затратах. Этот проект можно сохранить в Project Web App, но не публиковать. Убедитесь, что надстройка отображает данные с Project Server, но показывает **NA** для текущего проекта.

### <a name="procedure-6-test-the-add-in"></a>Процедура 6. Тестирование надстройки

1. Запустите Project профессиональный 2013, подключитесь к Project Web App и создайте тестовый проект. Назначьте задачи локальным ресурсам или ресурсам предприятия, настройте различные значения процента выполнения для некоторых задач и затем опубликуйте проект. Закройте Project, что позволит Visual Studio запустить Project для отладки надстройки.

2. В Visual Studio нажмите клавишу **F5**. Войдите в Project Web App и затем откройте проект, созданный на предыдущем шаге. Проект можно открыть в режиме чтения или в режиме редактирования.

3. На **вкладке PROJECT** ленты в раскрывающемся списке надстроек **Office** выберите **Hello ProjectData** (см. рис. 5). **Кнопка "Сравнить все проекты**" должна быть отключена.

    *Рис. 5. Запуск надстройки HelloProjectOData*

    ![Протестируйте приложение HelloProjectOData.](../images/pj15-hello-project-data-test-the-app.png)

4. В области **задач Hello ProjectData** выберите "Получить конечную точку **ProjectData"**. В **строке projectDataEndPoint** должен отображаться URL-адрес службы **ProjectData** , а также должна быть включена кнопка "Сравнить **все** проекты" (см. рис. 6).

5. Нажмите кнопку **Compare All Projects**. Надстройка может приостановить работу на время получения данных из службы **ProjectData**, а затем она должна отобразить отформатированные средние и текущие значения в таблице.

    *Рис. 6. Просмотр результатов запроса REST*

    ![Просмотр результатов запроса REST.](../images/pj15-hello-project-data-rest-results.png)

6. Проверьте выходные данные в текстовом поле. Они должны показывать путь к документу, запрос REST, сведения о состоянии и результаты JSON от вызовов **ajax** и **parseODataResult**. Выходные данные помогают понять, создать и отладить код в `parseODataResult` таком методе, как `projCost += Number(res.d.results[i].ProjectCost);`.

    Ниже приведен пример выходных данных с разрывами строк и пробелами, добавленными в текст для ясности для трех проектов в Project Web App экземпляре.

    ```json
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

7. Остановите отладку (нажмите **клавиши SHIFT+F5**), а затем нажмите **клавишу F5** еще раз, чтобы запустить новый экземпляр Project. В **диалоговом** окне входа выберите профиль локального  компьютера, а не Project Web App. Создайте или откройте MPP-файл локального проекта, откройте область задач **Hello ProjectData** и выберите "Получить конечную точку **ProjectData"**. Надстройка должна отобразить " **Нет подключения".** Ошибка  (см. рис. 7) и кнопка "Сравнить **все** проекты" должна оставаться отключенной.

   *Рис. 7. Использование надстройки без подключения к веб-приложению Project*

   ![Используйте приложение без подключения Project Web App подключения.](../images/pj15-hello-project-data-no-connection.png)

8. Остановите отладку и нажмите клавишу **F5** снова. Войдите в Project Web App и создайте проект, содержащий данные о материальных и трудовых затратах. Проект можно сохранить, но не публикуйте его.

   В области **задач Hello ProjectData** при выборе параметра "Сравнить все проекты" должно появиться синее значение **NA** для полей в  текущем столбце (см. рис. 8).

   *Рис. 8. Сравнение неопубликованного проекта с другими проектами*

   ![Сравните неопубликованный проект с другими пользователями.](../images/pj15-hello-project-data-not-published.png)

Даже если ваша надстройка работала правильно в предыдущих тестах, есть другие тесты, которые необходимо выполнить. Например:

- Откройте в Project Web App проект, который не содержит данных о материальных и трудовых затратах для задач. В полях столбца **Current (текущий)** должны отображаться нули.

- Протестируйте проект, не содержащий задачи.

- Если вы измените надстройку и опубликуете ее, необходимо запустить аналогичные тесты снова с опубликованной надстройкой. Другие вопросы см. в разделе [Дальнейшие действия](#next-steps).

> [!NOTE]
> Имеются ограничения на объем данных, который может быть возвращен в одном запросе службы **ProjectData**; этот объем данных меняется для разных сущностей. Например, набор сущностей `Projects` имеет ограничение по умолчанию в 100 проектов на запрос, `Risks` но набор сущностей имеет ограничение по умолчанию в 200. Для установки в рабочей среде код в примере **HelloProjectOData** необходимо изменить для поддержки запросов, содержащих более 100 проектов. Дополнительные сведения см. в [разделах "Дальнейшие действия"](#next-steps) и "Запрос веб-каналов [OData для данных отчетов Project"](/previous-versions/office/project-odata/jj163048(v=office.15)).

## <a name="example-code-for-the-helloprojectodata-add-in"></a>Пример кода для надстройки HelloProjectOData

### <a name="helloprojectodatahtml-file"></a>Файл HelloProjectOData.html

Приведенный ниже код находится в файле `Pages\HelloProjectOData.html` проекта **HelloProjectODataWeb**.

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file. -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files. -->
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

### <a name="helloprojectodatajs-file"></a>Файл HelloProjectOData.js

Приведенный ниже код находится в файле `Scripts\Office\HelloProjectOData.js` проекта **HelloProjectODataWeb**.

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

//  Functions to get and parse the Project Server reporting data./

// Get data about all projects on Project Server,
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project on the web.
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

### <a name="appcss-file"></a>Файл App.css

Приведенный ниже код находится в файле `Content\App.css` проекта **HelloProjectODataWeb**.

```css
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

### <a name="surfaceerrorsjs-file"></a>Файл SurfaceErrors.js

Вы можете скопировать код для файла SurfaceErrors.js из раздела _Надежное программирование_ статьи [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).

## <a name="next-steps"></a>Дальнейшие действия

Если **бы HelloProjectOData** была рабочей надстройки для продажи в AppSource или распространялись в каталоге приложений SharePoint, она была бы разработана по-другому. Например, здесь не было бы отладочных выходных данных в текстовом поле и, вероятно, не было бы кнопки для получения конечной точки **ProjectData**. Вам также потребуется переписать функцию `retireveOData` для обработки Project Web App экземпляров, которые содержат более 100 проектов.

Надстройка должна содержать дополнительные проверки ошибок, а также логику для записи, объяснения или демонстрации пограничных случаев. Например, если экземпляр Project Web App содержит 1000 проектов со средней продолжительностью в пять дней и средними затратами в $2400, а активный проект является единственным с продолжительностью более 20 дней, то сравнение материальных и трудовых затрат может быть перекошено. Это может быть показано с помощью частотной диаграммы. Вам необходимо добавить команды для отображения продолжительности, сравнения проектов с одинаковой продолжительностью или сравнения проектов из одного или разных отделов. Либо добавить возможность пользователю выбирать из списка полей, которые требуется отобразить.

Для других запросов службы **ProjectData** имеются ограничения на длину строки запроса, что влияет на число шагов, которые запрос может предпринять для выборки из родительской коллекции в объект в дочерней коллекции. Например, двухшаговый запрос **Projects** в **Tasks** для получения элементов задач работает, но трехшаговый запрос, такой как **Projects** в **Tasks** в **Assignments**, для получения элемента назначения может превысить максимальную длину URL-адреса по умолчанию. Дополнительные сведения см [. в разделе "Запрос веб-каналов OData для данных отчетов Project"](/previous-versions/office/project-odata/jj163048(v=office.15)).

Если вы изменяете **надстройку HelloProjectOData** для использования в рабочей области, выполните следующие действия.

- В файле HelloProjectOData.html для лучшей производительности измените ссылку office.js из локального проекта на ссылку CDN:

    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- Перепишите функцию `retrieveOData` , чтобы включить запросы для более чем 100 проектов. Например, можно получить число проектов с помощью запроса `~/ProjectData/Projects()/$count` и использовать оператор _$skip_ и оператор _$top_ в запросе REST для получения данных проекта. Запустите несколько запросов в цикле и затем усредните данные из всех запросов. Каждый запрос для данных проекта будет иметь следующий вид:

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`

  Дополнительные сведения см. в разделе параметров системных запросов [OData с использованием конечной точки REST](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps&preserve-view=true) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).

- Сведения о развертывании надстройки см. в статье [Публикация надстройки Office](../publish/publish.md).

## <a name="see-also"></a>См. также

- [Надстройки области задач для Project](project-add-ins.md)
- [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- [ProjectData — Справочник по службе Project OData](/previous-versions/office/project-odata/jj163015(v=office.15))
- [XML-манифест надстройки Office](../develop/add-in-manifests.md)
- [Публикация надстройки Office](../publish/publish.md)
