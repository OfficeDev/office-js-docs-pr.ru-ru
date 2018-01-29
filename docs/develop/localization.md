# <a name="localization-for-office-add-ins"></a>Локализация надстроек для Office

Вы можете реализовать любую схему локализации, которая подходит вашему Надстройка Office. API JavaScript и схема манифеста платформы Надстройки Office предоставляют несколько вариантов. Вы можете использовать API JavaScript для Office, чтобы определить языковые параметры и отображать строки, основываясь на языковых параметрах ведущего приложения, или интерпретировать и отображать данные на основе языковых параметров данных. Вы можете использовать манифест, чтобы указать расположение файла надстройки и описательной информации, зависящих от языковых параметров. Либо можно использовать сценарий Microsoft Ajax для поддержки глобализации и локализации.

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a>Определение параметров, зависящих от языка, с помощью API JavaScript

API JavaScript для Office предоставляет два свойства, которые поддерживают отображение и интерпретацию значений в соответствии с языковыми параметрами ведущего приложения и данными:

- [Context.displayLanguage][displayLanguage] задает языковой стандарт пользовательского интерфейса ведущего приложения. В следующем примере показано, как проверить, какой языковый параметр используется (en-US или fr-Fr), и отобразить приветствие на языке ведущего приложения.
    
    ```js
    function sayHelloWithDisplayLanguage() {
        var myLanguage = Office.context.displayLanguage;
        switch (myLanguage) {
            case 'en-US':
                write('Hello!');
                break;
            case 'fr-FR':
                write('Bonjour!');
                break;
        }
    }
    
    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message; 
    }
    ```

- [Context.contentLanguage][contentLanguage] задает языковой стандарт данных. Вы можете не проверять свойство [displayLanguage], а назначить свойству [contentLanguage] значение `myLanguage` и воспользоваться тем же кодом для отображения приветствия на языке данных:
    
    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a>Управление локализацией через манифест


Каждая надстройка Office задает в своем манифесте элемент [DefaultLocale] и языковой стандарт. По умолчанию платформа надстроек Office и ведущие приложения Office применяют значения элементов [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl] и [SourceLocation] ко всем языковым стандартам. Чтобы изменить значения для определенных языковых стандартов, укажите для любого из этих пяти элементов дочерний элемент [Override]. Значение элемента [DefaultLocale] и атрибута `Locale` элемента [Override] указывается в соответствии со спецификацией [RFC 3066], "Теги для идентификации языков". В таблице 1 описана поддержка локализации для этих элементов.

**Таблица 1. Поддержка локализации**


|**Элемент**|**Поддержка локализации**|
|:-----|:-----|
|[Описание]   |В рамках каждого указанного вами языкового стандарта для пользователей отображается локализованное описание надстройки в Магазине Office (или в частном каталоге).<br/>В случае надстроек Outlook пользователи могут видеть описание в Центре администрирования Exchange после установки.|
|[DisplayName]   |В рамках каждого указанного вами языкового стандарта для пользователей отображается локализованное описание надстройки в Магазине Office (или в частном каталоге).<br/>В случае надстроек Outlook пользователи могут видеть отображаемое имя в качестве метки для кнопки надстройки Outlook и в Центре администрирования Exchange после установки.<br/>В случае контентных надстроек и надстроек области задач пользователи могут видеть отображаемое имя на ленте после установки надстройки.|
|[IconUrl]        |Изображение значка является необязательным. Можно использовать ту же методику переопределений, чтобы задать определенное изображение для определенной культуры. Если вы используете значок и локализуете его, пользователи с заданными языковыми параметрами могут видеть локализованный значок надстройки.<br/>В случае надстроек Outlook пользователи могут видеть значок в Центре администрирования Exchange после установки надстройки.<br/>В случае контентных надстроек и надстроек области задач пользователи могут видеть значок на ленте после установки надстройки.|
|[HighResolutionIconUrl] <br/><br/>**Важно!** Этот элемент доступен только для надстроек, использующих схему манифеста версии 1.1.|Изображение значка с высоким разрешением не является обязательным, но если оно указано, то должно находиться после элемента [IconUrl]. Если указан параметр [HighResolutionIconUrl] и надстройка установлена на устройстве, поддерживающем высокое разрешение, то вместо значения [IconUrl] используется значение [HighResolutionIconUrl].<br/>Можно использовать ту же методику переопределений, чтобы задать определенное изображение для определенной культуры. Если вы используете значок и локализуете его, пользователи с заданными языковыми параметрами могут видеть локализованный значок надстройки.<br/>В случае надстроек Outlook пользователи могут видеть значок в Центре администрирования Exchange после установки надстройки.<br/>При работе с контентными надстройками и надстройками области задач после установки надстройки для пользователей отображается значок на ленте.|
|[Resources] <br/><br/>**Важно!** Этот элемент доступен только для надстроек, в которых используется схема манифеста версии 1.1.   |Для пользователей в каждой указываемой вами локали отображаются ресурсы строк и значков, которые вы специально создаете для надстройки в этой локали. |
|[SourceLocation]   |Пользователи, имеющие указанные языковые параметры, могут видеть веб-страницу, специально разработанную для надстройки с заданными языковыми параметрами. |


 > **Примечание.** Вы можете локализовать описание и отображаемое имя только для языковых стандартов, которые поддерживаются в Office. Список языков и языковых стандартов для текущего выпуска Office см. в статье [Идентификаторы языков и значения OptionState Id в Office 2013](http://technet.microsoft.com/ru-ru/library/cc179219.aspx).


### <a name="examples"></a>Примеры

Например, Надстройка Office может задать [DefaultLocale] как `en-us`. Для элемента [DisplayName] надстройка может задать дочерний элемент [Override] для языкового параметра `fr-fr`, как показано ниже. 


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

 > **Примечание.** Если вам необходимо локализовать несколько областей в семействе языков, например `de-de` и `de-at`, рекомендуется разделить элементы `Override` для каждой области. Использование только названия языка, в данном случае `de`, не поддерживается ни для одного сочетания ведущих приложений Office и платформ.

Это значит, что по умолчанию надстройка использует языковой стандарт `en-us`. Пользователи видят отображаемое имя Video player (видеопроигрыватель) на английском языке для всех языковых стандартов за исключением случаев, когда на клиентском компьютере используется языковой стандарт `fr-fr`. В этом случае пользователи увидят отображаемое имя Lecteur video на французском языке.

> **Примечание.** Вы можете указать только одно переопределение на язык, в том числе для языкового стандарта по умолчанию. Например, если по умолчанию используется языковой стандарт `en-us`, нельзя также указать переопределение для `en-us`. 

В следующем примере применяется переопределение языкового параметра для элемента [Description]. Сначала он задает языковой параметр по умолчанию `en-us` и описание на английском языке, а затем указывает оператор [Override] с описанием на французском языке для языкового параметра `fr-fr`:

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive 
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook et Outlook Web App."/>
</Description>
```

Это значит, что надстройка предполагает языковой стандарт `en-us` по умолчанию. Пользователи увидят описание на английском языке в атрибуте `DefaultValue` для всех языковых стандартов, если на клиентском компьютере не выбран языковой стандарт `fr-fr`. В этом случае они увидят описание на французском языке.

В следующем примере надстройка задает отдельное приложение, которое больше подходит для языкового параметра и культуры `fr-fr`. Пользователи видят изображение DefaultLogo.png по умолчанию, кроме тех случаев, когда языковой параметр клиентского компьютера равен `fr-fr`. В этом случае пользователи видят изображение FrenchLogo.png. 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
    <Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

В примере ниже показано, как локализовать ресурс в разделе `Resources`. Здесь применяется переопределение локали для изображения, и используется изображение, более подходящее для языка и региональных параметров `ja-jp`.

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


В случае элемента [SourceLocation] поддержка дополнительных языковых параметров означает предоставление отдельного исходного HTML-файла для каждого из указанных языковых параметров. Пользователи с заданными языковыми параметрами могут видеть настраиваемые для них веб-страницы.

В случае надстроек Outlook элемент [SourceLocation] также сопоставляется с форм-фактором. Это позволяет предоставлять отдельный локализованный исходный HTML-файл для каждого соответствующего форм-фактора. Можно задать один или несколько дочерних элементов [Override] в каждом необходимом элементе параметров ([DesktopSettings], [TabletSettings] или [PhoneSettings]). В следующем примере показаны элементы параметров для форм-факторов настольного компьютера, планшета и смартфона, каждый из которых имеет один HTML-файл для языкового параметра по умолчанию и другой файл для французского языкового параметра.


```xml
<DesktopSettings>
   <SourceLocation DefaultValue="https://contoso.com/Desktop.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Desktop.html" />
   </SourceLocation>
   <RequestedHeight>250</RequestedHeight>
</DesktopSettings>
<TabletSettings>
   <SourceLocation DefaultValue="https://contoso.com/Tablet.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Tablet.html" />
   </SourceLocation>
   <RequestedHeight>200</RequestedHeight>
</TabletSettings>
<PhoneSettings>
   <SourceLocation DefaultValue="https://contoso.com/Mobile.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Mobile.html" />
   </SourceLocation>

</PhoneSettings>

```

## <a name="match-datetime-format-with-client-locale"></a>Приведение формата даты и времени в соответствие языковому стандарту клиента

Получить языковой стандарт интерфейса ведущего приложения можно с помощью свойства [displayLanguage]. Затем можно отобразить значения даты и времени в формате, соответствующем текущему языковому стандарту ведущего приложения. Для этого можно подготовить файл ресурсов и указать в нем формат даты и времени для каждого языкового стандарта, который поддерживает надстройка Office. После этого надстройка сможет сопоставлять соответствующий формат даты и времени с языковым стандартом, полученным из свойства [displayLanguage], во время выполнения.

Получить языковой стандарт данных ведущего приложения можно с помощью свойства [contentLanguage]. На основе этого значения можно интерпретировать или отображать строки даты и времени. Например, в языковом стандарте `jp-JP` дата и время выражаются так: `yyyy/MM/dd`, а в языковом стандарте `fr-FR` так: `dd/MM/yyyy`.


## <a name="use-ajax-for-globalization-and-localization"></a>Использование Ajax для глобализации и локализации


Если для создания Надстройки Office вы используете Visual Studio, платформа .NET Framework и Ajax предоставляют способы глобализации и локализации файлов клиентских скриптов.

Можно глобализировать и использовать расширения типов JavaScript [Date](http://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) и [Number](http://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) и объект JavaScript [Date](http://msdn.microsoft.com/library/ce2202bb-7ec9-4f5a-bf48-3a04feff283e.aspx) в коде JavaScript для Надстройка Office, чтобы отображать значения в зависимости от языковых параметров, заданных в текущем браузере. Дополнительные сведения см. в статье [Walkthrough: Globalizing a Date by Using Client Script](http://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).

Можно включить локализованные строки ресурсов напрямую в отдельные файлы JavaScript, чтобы предоставить клиентские файлы скриптов для разных языковых параметров, задаваемых в браузере или предоставляемых пользователем. Создайте отдельный файл скрипта для каждого поддерживаемого языкового параметра. В каждый файл скрипта включите объект в формате JSON, содержащий строки ресурсов для соответствующего языкового параметра. Локализованные значения применяются во время выполнения скрипта в браузере. 


## <a name="example-build-a-localized-office-add-in"></a>Пример. Создание локализованной надстройки Office

В этом разделе представлены примеры того, как локализовать описание, отображаемое имя и пользовательский интерфейс Надстройка Office.

Чтобы запустить предоставленный образец кода, настройте Microsoft Office 2013 на своем компьютере для использования дополнительных языков, чтобы надстройку можно было тестировать на переключение языка меню и команд, для редактирования и/или проверки.

Кроме того, необходимо создать приложение Visual Studio 2015 для проекта надстройки Office.

 > **Примечание.** Скачать Visual Studio 2015 можно на [странице Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs). На данной странице также имеется ссылка на Office Developer Tools.

### <a name="configure-office-2013-to-use-additional-languages-for-display-or-editing"></a>Настройка Office 2013 на использование дополнительных языков для отображения или редактирования

Для установки дополнительного языка можно использовать языковой пакет Office 2013. Дополнительные сведения о языковых пакетах и способах их получения см. в статье [Языковые возможности Office 2013](http://office.microsoft.com/en-us/language-packs/).

 > **Примечание.** Если вы подписчик MSDN, то для вас уже могут быть доступны языковые пакеты Office 2013. Чтобы определить, позволяет ли ваша подписка скачать языковые пакеты Office 2013, перейдите на [главную страницу подписок на MSDN](https://msdn.microsoft.com/subscriptions/manage/), в поле **Загрузка программного обеспечения** введите "языковой пакет Office 2013", щелкните **Поиск**, а затем выберите **Продукты, доступные для моей подписки**. В разделе **Язык** установите флажок нужного языкового пакета и щелкните **Перейти**. 

После установки языкового пакета вы можете настроить Office 2013 на использование установленного языка для пользовательского интерфейса и/или для редактирования содержимого документов. В примере в этой статье используется установка Office 2013, в которой применяется испанский языковой пакет.

### <a name="create-an-office-add-in-project"></a>Создание проекта надстройки Office

1. В Visual Studio выберите команду **Файл** > **Создать проект**.
    
2. В области **Шаблоны** диалогового окна **Новый проект** разверните узел **Visual Basic** или **Visual C#**, узел **Office/SharePoint**, а затем выберите **Надстройки Office**.
    
3. Выберите **Надстройка Office** и присвойте надстройке имя, например WorldReadyAddIn. Нажмите кнопку **ОК**.
    
4. В диалоговом окне **Создание надстройки Office** выберите **Надстройка области задач** и нажмите кнопку **Далее**. На следующей странице снимите флажки для всех ведущих приложений, кроме Word. Выберите **Готово** для создания проекта.
    

### <a name="localize-the-text-used-in-your-add-in"></a>Локализация текста, используемого в вашей надстройке

Текст, который нужно локализовать на другом языке, отображается в двух областях:

-  **Отображаемое имя и описание надстройки**. Они управляются записями в файле манифеста приложения.
    
-  **Пользовательский интерфейс надстройки**. Вы можете локализовать строки, отображаемые в пользовательском интерфейсе, с помощью кода JavaScript, например с помощью отдельного файла ресурсов с локализованными строками.
    
Локализация отображаемого имени и описания надстройки:

1. В **обозревателе решений** разверните узлы **WorldReadyAddIn** и **WorldReadyAddInManifest**, а затем выберите **WorldReadyAddIn.xml**.
    
2. В файле WorldReadyAddInManifest.xml замените элементы [DisplayName] и [Description] приведенным ниже блоком кода.
    
     > **Примечание.** Вы можете заменить локализованные строки на испанском языке, используемые в этом примере для элементов [DisplayName] и [Description], локализованными строками на любом другом языке.

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. После изменения отображаемого языка для Office 2013, к примеру, с английского на испанский и последующего запуска надстройки отображаемое имя и описание надстройки локализуются. 
    
Настройка пользовательского интерфейса надстройки:

1. В **обозревателе решений**Visual Studio выберите элемент **Home.html**.
    
2. Замените HTML-код в файле Home.html на следующий код.
    
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.8.2.js" type="text/javascript"></script>
    
        <link href="../../Content/Office.css" rel="stylesheet" type="text/css" />
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    
        <!-- To enable offline debugging using a local reference to Office.js, use:                        -->
        <!-- <script src="../../Scripts/Office/MicrosoftAjax.js" type="text/javascript"></script>          -->
        <!--    <script src="../../Scripts/Office/1.0/office.js" type="text/javascript"></script>          -->
    
        <link href="../App.css" rel="stylesheet" type="text/css" />
        <script src="../App.js" type="text/javascript"></script>
    
        <link href="Home.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script> <body>
        <!-- Page content -->
        <div id="content-header">
            <div class="padding">
                <h1 id="greeting"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div>
                    <p id="about"></p>
                </div>            
            </div>
        </div>
    </head>
    </html>
    ```

3. В Visual Studio откройте вкладку **Файл** и выберите команду **Сохранить AddIn\Home\Home.html**.
    
На рисунке 3 показаны элемент заголовка (h1) и элемент абзаца (p), в которых будет отображаться локализованный текст при выполнении примера надстройки.

**Рис. 3. Интерфейс надстройки**

![Пользовательский интерфейс приложения с выделенными разделами.](../images/off15App_HowToLocalize_fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a>Добавление файла ресурсов с локализованными строками

Файл ресурсов JavaScript содержит строки, используемые для пользовательского интерфейса надстройки. Пользовательский интерфейс в примере надстройки содержит элемент h1, отображающий приветствие, и элемент p, который знакомит пользователя с надстройкой. 

Чтобы включить локализованные строки для заголовка и абзаца, нужно поместить строки в отдельный файл ресурса. Файл ресурса создает объект JavaScript, который содержит отдельный объект Нотация объектов JavaScript (JSON) для каждого набора локализованных строк. Файл ресурса также предоставляет метод для получения соответствующего объекта JSON для определенного региона. 

Добавление файла ресурса в проект надстройки:

1. В **обозревателе решений**Visual Studio выберите папку **Add-in** в веб-проекте образца приложения и выберите команды **Добавить**  >  **Файл JavaScript**.
    
2. В диалоговом окне **Указать имя для элемента** введите UIStrings.js.
    
3. Добавьте в файл UIStrings.js следующий код.

    ```js
    /* Store the locale-specific strings */
    
    var UIStrings = (function ()
    {
        "use strict";
    
        var UIStrings = {};
    
        // JSON object for English strings
        UIStrings.EN =
        {        
            "Greeting": "Welcome",
            "Introduction": "This is my localized add-in."        
        };
    
    
        // JSON object for Spanish strings
        UIStrings.ES =
        {        
            "Greeting": "Bienvenido",
            "Introduction": "Esta es mi aplicación localizada."
        };
    
        UIStrings.getLocaleStrings = function (locale)
        {
            var text;
            
            // Get the resource strings that match the language.
            switch (locale)
            {
                case 'en-US':
                    text = UIStrings.EN;
                    break;
                case 'es-ES':
                    text = UIStrings.ES;
                    break;
                default:
                    text = UIStrings.EN;
                    break;
            }
    
            return text;
        };
    
        return UIStrings;
    })();
    ```

Файл ресурса UIStrings.js создает объект **UIStrings**, который содержит локализованные строки пользовательского интерфейса надстройки. 

### <a name="localize-the-text-used-for-the-add-in-ui"></a>Локализация текста, используемого для пользовательского интерфейса надстройки

Чтобы использовать в надстройке файл ресурсов, вам потребуется добавить для него тег сценария в файл Home.html. При загрузке Home.html выполняется UIStrings.js, и объект **UIStrings**, используемый для получения строк, становится доступен в коде. Добавьте следующий HTML-код в тег заголовка для Home.html, чтобы сделать **UIStrings** доступным в коде.

```html
<!-- Resource file for localized strings:                                                          -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

Теперь вы можете использовать объект **UIStrings**, чтобы задать строки для пользовательского интерфейса надстройки.

Если вы хотите изменить локализацию в зависимости от языка, используемого для меню и команд в ведущем приложении, используйте свойство **Office.context.displayLanguage**, чтобы получить языковой стандарт для этого языка. Например, если в ведущем приложении для отображения меню и команд используется испанский язык, свойство **Office.context.displayLanguage** возвращает код языка es-ES.

Если вы хотите изменить локализацию в зависимости от языка, используемого для редактирования содержимого документов, используйте свойство **Office.context.contentLanguage**, чтобы получить языковой стандарт для этого языка. Например, если в ведущем приложении для редактирования содержимого документов используется испанский язык, свойство **Office.context.contentLanguage** возвращает код языка es-ES.

Узнав, какой язык использует ведущее приложение, вы можете использовать **UIStrings**, чтобы получить набор локализованных строк на языке ведущего приложения.

Замените код в файле Home.js на следующий код. Этот код показывает, как вы можете изменить строки, используемые в элементах пользовательского интерфейса из файла Home.html, в зависимости от языка отображаемых элементов или языка редактирования в ведущем приложении.

 > **Примечание.** Для переключения локализаций надстройки на основании языка, используемого для редактирования, раскомментируйте строку кода `var myLanguage = Office.context.contentLanguage;` и закомментируйте строку кода `var myLanguage = Office.context.displayLanguage;`.

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {
       
        $(document).ready(function () {
            app.initialize();

            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // var myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the host application.
            var myLanguage = Office.context.displayLanguage;            
            var UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);            

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Instruction);
        });
    };    
})();
```

### <a name="test-your-localized-add-in"></a>Тестирование локализованной надстройки

Чтобы тестировать локализованную надстройку, измените язык, используемый для отображения или редактирования в ведущем приложении, а затем запустите надстройку. 

Изменение языка, используемого для отображения или редактирования в надстройке:

1. В Word 2013 выберите команды **Файл**, **Параметры**, **Язык**. На рисунке 4 показано диалоговое окно **Параметры Word**, открытое на вкладке языка.
    
    **Рис. 4. Параметры языка в диалоговом окне "Параметры" Word 2013**

    ![Диалоговое окно "Параметры" Word 2013.](../images/off15App_HowToLocalize_fig04.png)

2. В области **Выбор языков интерфейса и справки** выберите язык, на котором должны отображаться данные (например, испанский), а затем нажмите стрелку вверх, чтобы переместить испанский язык в начало списка. Вы также можете изменить язык, используемый для редактирования. В области **Выбор языков редактирования** выберите язык, который хотите использовать для редактирования (например, испанский), а затем нажмите кнопку **Использовать по умолчанию**.
    
3. Нажмите кнопку **ОК**, чтобы подтвердить выбор, а затем закройте Word.
    
Запустите образец надстройки. Надстройка области задач загружается в Word 2013, а строки в пользовательском интерфейсе надстройки меняются в соответствии с языком ведущего приложения, как показано на рис. 5.


**Рис. 5. Пользовательский интерфейс надстройки с локализованным текстом**

![Приложение с локализованным текстом пользовательского интерфейса.](../images/off15App_HowToLocalize_fig05.png)

## <a name="additional-resources"></a>Дополнительные ресурсы

- [Рекомендации по проектированию надстроек Office](../design/add-in-design.md)
    
- [Идентификаторы языков и значения OptionState Id в Office 2013](http://technet.microsoft.com/ru-ru/library/cc179219%28Office.15%29.aspx)

[DefaultLocale]:         http://dev.office.com/reference/add-ins/manifest/defaultlocale
[Описание]:           http://dev.office.com/reference/add-ins/manifest/description
[DisplayName]:           http://dev.office.com/reference/add-ins/manifest/displayname
[IconUrl]:               http://dev.office.com/reference/add-ins/manifest/iconurl
[HighResolutionIconUrl]: http://dev.office.com/reference/add-ins/manifest/highresolutioniconurl
[Resources]:             https://dev.office.com/reference/add-ins/manifest/resources
[SourceLocation]:        http://dev.office.com/reference/add-ins/manifest/sourcelocation
[Переопределение]:              http://dev.office.com/reference/add-ins/manifest/override
[DesktopSettings]:       http://dev.office.com/reference/add-ins/manifest/desktopsettings
[TabletSettings]:        http://dev.office.com/reference/add-ins/manifest/tabletsettings
[PhoneSettings]:         http://dev.office.com/reference/add-ins/manifest/phonesettings
[displayLanguage]:  http://dev.office.com/reference/add-ins/shared/office.context.displaylanguage 
[contentLanguage]:  http://dev.office.com/reference/add-ins/shared/office.context.contentlanguage 
[RFC 3066]: https://www.rfc-editor.org/info/rfc3066