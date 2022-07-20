---
title: Локализация надстроек для Office
description: Используйте API JavaScript для Office для определения языкового стандарта и отображения строк на основе языкового стандарта приложения Office, а также для интерпретации или отображения данных на основе языкового стандарта данных.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7f80f48c1c933ac6ef7c2e37fb3efcf3dd7ae073
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889466"
---
# <a name="localization-for-office-add-ins"></a>Локализация надстроек для Office

Вы можете реализовать любую схему локализации, которая подходит вашему Надстройка Office. API JavaScript и схема манифеста платформы Надстройки Office предоставляют несколько вариантов. API JavaScript для Office можно использовать для определения языкового стандарта и отображения строк на основе языкового стандарта приложения Office, а также для интерпретации или отображения данных на основе языкового стандарта данных. Вы можете использовать манифест, чтобы указать расположение файла надстройки и описательной информации, зависящих от языковых параметров. Либо можно использовать сценарий Microsoft Ajax для поддержки глобализации и локализации.

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a>Определение параметров, зависящих от языка, с помощью API JavaScript

API JavaScript для Office предоставляет два свойства, которые поддерживают отображение или интерпретацию значений, согласованных с языковым стандартом приложения и данных Office.

- [Context.displayLanguage][displayLanguage] задает языковой стандарт (или язык) пользовательского интерфейса приложения Office. В следующем примере проверяется, использует ли приложение Office языковой стандарт en-US или fr-FR и отображает приветствие для конкретного языкового стандарта.

    ```js
    function sayHelloWithDisplayLanguage() {
        const myLanguage = Office.context.displayLanguage;
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

- [Context.contentLanguage][contentLanguage] задает языковой стандарт данных. Чтобы расширить последний пример кода, вместо проверки свойства [displayLanguage] назначьте `myLanguage` значение свойства [contentLanguage] и используйте остальную часть того же кода для отображения приветствия на основе языкового стандарта данных.

    ```js
    const myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a>Управление локализацией через манифест

Каждое Надстройка Office задает в своем манифесте элемент [DefaultLocale] и языковой параметр. По умолчанию платформа надстроек Office и клиентские приложения Office применяют значения элементов [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl] и [SourceLocation] ко всем языковым стандартам. Чтобы изменить значения для определенных языковых стандартов, укажите для любого из этих пяти элементов дочерний элемент [Override]. Значение элемента [DefaultLocale] и атрибута `Locale` элемента [Override] указывается в соответствии со спецификацией [RFC 3066], "Теги для идентификации языков". В таблице 1 описана поддержка локализации для этих элементов.

*Таблица 1. Поддержка локализации*

|**Элемент**|**Поддержка локализации**|
|:-----|:-----|
|[Описание]   |Для каждого заданного языкового стандарта пользователи могут видеть локализованное описание надстройки в AppSource (или частном каталоге).<br/>В случае надстроек Outlook пользователи смогут увидеть описание в Центре администрирования Exchange после установки.|
|[DisplayName]   |Для каждого заданного языкового стандарта пользователи могут видеть локализованное описание надстройки в AppSource (или частном каталоге).<br/>В случае надстроек Outlook пользователи смогут увидеть отображаемое имя в качестве метки для кнопки надстройки Outlook и в Центре администрирования Exchange после установки.<br/>В случае контентных надстроек и надстроек области задач пользователи могут видеть отображаемое имя на ленте после установки надстройки.|
|[IconUrl]        |Изображение значка является необязательным. Можно использовать ту же методику переопределений, чтобы задать определенное изображение для определенной культуры. Если вы используете значок и локализуете его, пользователи с заданными языковыми параметрами могут видеть локализованный значок надстройки.<br/>В случае надстроек Outlook пользователи могут видеть значок в Центре администрирования Exchange после установки надстройки.<br/>После установки надстроек области задач и контентных надстроек пользователи видят значок на ленте.|
|[HighResolutionIconUrl] **Важно!** Этот элемент доступен только для надстроек, использующих схему манифеста версии 1.1.|Изображение значка с высоким разрешением не является обязательным, но если оно указано, то должно находиться после элемента [IconUrl]. Если указан параметр [HighResolutionIconUrl] и надстройка установлена на устройстве, поддерживающем высокое разрешение, то вместо значения [IconUrl] используется значение [HighResolutionIconUrl].<br/>Можно использовать ту же методику переопределений, чтобы задать определенное изображение для определенной культуры. Если вы используете значок и локализуете его, пользователи с заданными языковыми параметрами могут видеть локализованный значок надстройки.<br/>В случае надстроек Outlook пользователи могут видеть значок в Центре администрирования Exchange после установки надстройки.<br/>После установки надстроек области задач и контентных надстроек пользователи видят значок на ленте.|
|[Resources] **Важно!** Этот элемент доступен только для надстроек, в которых используется схема манифеста версии 1.1.   |Для пользователей в каждой указываемой вами локали отображаются ресурсы строк и значков, которые вы специально создаете для надстройки в этой локали. |
|[SourceLocation]   |Пользователи каждого языкового стандарта видят веб-страницу, специально разработанную для использования надстройки с этим стандартом. |

> [!NOTE]
> Локализовать описание и отображаемое имя можно только для языковых стандартов, которые поддерживаются в Office. Список языков и языковых стандартов для текущего выпуска Office см. в статье [Идентификаторы языков и значения OptionState Id в Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)).

### <a name="examples"></a>Примеры

Например, надстройка Office может задать для параметра [DefaultLocale] значения `en-us`. Для элемента [DisplayName] надстройка может задать дочерний элемент [Override], соответствующий языковому стандарту `fr-fr`, как показано ниже.

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> Если вам необходимо локализовать несколько областей в семействе языков, например `de-de` и `de-at`, рекомендуется разделить элементы `Override` для каждой области. Использование только имени языка в `de`данном случае поддерживается не во всех сочетаниях клиентских приложений и платформ Office.

Это значит, что по умолчанию надстройка использует языковой стандарт `en-us`. Пользователи видят отображаемое имя Video player (видеопроигрыватель) на английском языке для всех языковых стандартов за исключением случаев, когда на клиентском компьютере используется языковой стандарт `fr-fr`. В этом случае пользователи увидят отображаемое имя Lecteur video на французском языке.

> [!NOTE]
> Вы можете указать только одно переопределение на язык, в том числе для языкового стандарта по умолчанию. Например, если по умолчанию используется языковой стандарт `en-us`, невозможно также указать переопределение для `en-us`.

В следующем примере применяется переопределение языкового стандарта для элемента [Description] . Сначала он задает языковой `en-us` стандарт по умолчанию и описание на английском языке, а затем указывает инструкцию [Override] `fr-fr` с описанием языкового стандарта на французском языке.

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook."/>
</Description>
```

Это значит, что надстройка предполагает языковой стандарт `en-us` по умолчанию. Пользователи увидят описание на английском языке в атрибуте `DefaultValue` для всех языковых стандартов, если на клиентском компьютере не выбран языковой стандарт `fr-fr`. В этом случае они увидят описание на французском языке.

В следующем примере надстройка задает отдельное приложение, которое больше подходит для языкового стандарта и региональных параметров `fr-fr`. Пользователи видят изображение DefaultLogo.png по умолчанию, кроме тех случаев, когда на клиентском компьютере используется языковой стандарт `fr-fr`. В этом случае пользователи видят изображение FrenchLogo.png.

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

В случае элемента [SourceLocation] поддержка дополнительных языковых стандартов означает предоставление отдельного исходного HTML-файла для каждого из указанных языковых стандартов. Пользователи с заданными языковыми стандартами увидят настраиваемые для них веб-страницы.

В случае надстроек Outlook элемент [SourceLocation] также сопоставляется с форм-фактором. Это позволяет предоставлять отдельный локализованный исходный HTML-файл для каждого соответствующего форм-фактора. Вы можете задать один или несколько дочерних элементов [Override] в каждом применимом элементе параметров ([DesktopSettings], [TabletSettings] или [PhoneSettings]). В приведенном ниже примере показаны элементы параметров для форм-факторов настольного компьютера, планшета и смартфона, каждому из которых соответствует один HTML-файл для языкового стандарта по умолчанию и другой файл для французского языкового стандарта.

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

## <a name="localize-extended-overrides"></a>Локализация расширенных переопределений

Некоторые функции расширяемости надстроек Office, такие как сочетания клавиш, настраиваются с помощью JSON-файлов, размещенных на сервере, а не XML-манифеста надстройки. В этом разделе предполагается, что вы знакомы с расширенными переопределениями. См [. раздел "Работа с расширенными переопределениями элемента манифеста](extended-overrides.md) и [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) ".

Используйте атрибут `ResourceUrl` элемента [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) , чтобы указать Office на файл локализованных ресурсов. Ниже приведен пример.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

В файле расширенных переопределений вместо строк используются маркеры. Строки имен маркеров в файле ресурсов. Ниже приведен пример назначения сочетания клавиш функции (определенной в другом месте), отображаемой в области задач надстройки. Обратите внимание на эту разметку:

- Пример не является вполне допустимым. (Ниже мы добавим необходимое дополнительное свойство.)
- Маркеры должны иметь формат **${resource.*имя ресурса*}**.

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ] 
}
```

Файл ресурсов, который также имеет формат JSON, имеет свойство верхнего уровня, `resources` которое делится на вложенные свойства по языковой стандарту. Для каждого языкового стандарта строка назначается каждому токену, который использовался в файле расширенных переопределений. Ниже приведен пример, в котором есть строки для и `en-us` `fr-fr`. В этом примере сочетание клавиш одинаково в обоих языковых стандартах, но это не всегда так, особенно при локализации языковых стандартов, которые имеют другой алфавит или систему письма, и, следовательно, другую клавиатуру.

```json
{
    "resources":{ 
        "en-us": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            }, 
        },
        "fr-fr": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Afficher le volet de tâche pour add-in",
              } 
        }
    }
}
```

В файле нет `default` свойства, которое является одноранговым для разделов `en-us` и разделов `fr-fr` . Это связано с тем, что строки по умолчанию, используемые, если языковой стандарт ведущего приложения Office не соответствует ни одному из свойств *ll-cc* в файле ресурсов, должны быть определены в самом файле расширенных переопределений. Определение строк по умолчанию непосредственно в расширенном файле переопределений гарантирует, что Office не скачивает файл ресурсов, если языковой стандарт приложения Office соответствует стандарту по умолчанию надстройки (как указано в манифесте). Ниже приведена исправленная версия предыдущего примера файла расширенных переопределений, в котором используются маркеры ресурсов.

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ],
    "resources": { 
        "default": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            } 
        }
    }
}
```

## <a name="match-datetime-format-with-client-locale"></a>Приведение формата даты и времени к языковым параметрам клиента

Языковой стандарт пользовательского интерфейса клиентского приложения Office можно получить с помощью свойства **[displayLanguage]** . Затем можно отобразить значения даты и времени в формате, согласованном с текущим языковым стандартом приложения Office. Один из способов сделать это — подготовить файл ресурсов, в котором задан формат отображения даты и времени для использования с каждым из языковых параметров, поддерживаемых Надстройка Office. Во время выполнения надстройка может использовать файл ресурсов и сопоставлять соответствующий формат даты и времени с языковым стандартом, полученным из свойства **[displayLanguage]** .

Языковой стандарт данных клиентского приложения Office можно получить с помощью свойства [contentLanguage] . На основе этого значения можно интерпретировать или отображать строки даты и времени. Например, в языковом стандарте `jp-JP` дата и время выражаются так: `yyyy/MM/dd`, а в языковом стандарте `fr-FR` так: `dd/MM/yyyy`.

## <a name="use-ajax-for-globalization-and-localization"></a>Использование Ajax для глобализации и локализации

Если для создания Надстройки Office вы используете Visual Studio, платформа .NET Framework и Ajax предоставляют способы глобализации и локализации файлов клиентских скриптов.

Можно глобализировать и использовать расширения типов JavaScript [Date](/previous-versions/bb310850(v=vs.140)) и [Number](/previous-versions/bb310835(v=vs.140)) и объект JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) в коде JavaScript для Надстройка Office, чтобы отображать значения в зависимости от языковых параметров, заданных в текущем браузере. Дополнительные сведения см. в статье [Walkthrough: Globalizing a Date by Using Client Script](/previous-versions/bb386581(v=vs.140)).

Можно включить локализованные строки ресурсов напрямую в отдельные файлы JavaScript, чтобы предоставить клиентские файлы скриптов для разных языковых параметров, задаваемых в браузере или предоставляемых пользователем. Создайте отдельный файл скрипта для каждого поддерживаемого языкового параметра. В каждый файл скрипта включите объект в формате JSON, содержащий строки ресурсов для соответствующего языкового параметра. Локализованные значения применяются во время выполнения скрипта в браузере.

## <a name="example-build-a-localized-office-add-in"></a>Пример. Создание локализованной надстройки Office

В этом разделе представлены примеры того, как локализовать описание, отображаемое имя и пользовательский интерфейс Надстройка Office.

> [!NOTE]
> Чтобы скачать Visual Studio 2019, перейдите на страницу [интегрированной среды разработки Visual Studio](https://visualstudio.microsoft.com/vs/). Во время установки потребуется выбрать рабочую нагрузку разработки для Office и SharePoint.

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a>Настройка Office на использование дополнительных языков для отображения или редактирования

Чтобы запустить предоставленный пример кода, настройте Office на компьютере для использования дополнительных языков, чтобы можно было протестировать надстройку, переключив язык, используемый для отображения в меню и командах, для редактирования и проверки правописания, или и то, и другое.

Для установки дополнительного языка можно использовать языковой пакет Office. Дополнительные сведения о языковых пакетах и способах их получения см. на странице [дополнительных языковых пакетов для Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).

После установки языкового пакета вы можете настроить Office на использование установленного языка для пользовательского интерфейса и/или для редактирования содержимого документов. В примере в этой статье используется установка Office, в которой применяется испанский языковой пакет.

### <a name="create-an-office-add-in-project"></a>Создание проекта надстройки Office

Вам потребуется создать проект надстройки Office для Visual Studio 2019.

> [!NOTE]
> Если вы еще не установили Visual Studio 2019, инструкции по скачии см. на странице интегрированной среды разработки [Visual Studio](https://visualstudio.microsoft.com/vs/) . Во время установки потребуется выбрать рабочую нагрузку разработки Office и SharePoint. Если вы ранее установили Visual Studio 2019[, используйте](/visualstudio/install/modify-visual-studio/) Visual Studio Installer, чтобы убедиться, что рабочая нагрузка разработки Office или SharePoint установлена.

1. Выберите **Создание нового проекта**.

1. Используя поле поиска, введите **надстройка**. Выберите вариант **Веб-надстройка Word** и нажмите кнопку **Далее**.

1. Приведите **проекту имя WorldReadyAddIn** и нажмите кнопку **"Создать"**.

1. Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.

### <a name="localize-the-text-used-in-your-add-in"></a>Локализация текста, используемого в вашей надстройке

Текст, который требуется локализовать для другого языка, отображается в двух областях.

- **Отображаемое имя и описание надстройки**. Они управляются записями в файле манифеста приложения.

- **Пользовательский интерфейс надстройки**. Вы можете локализовать строки, отображаемые в пользовательском интерфейсе надстройки, с помощью кода JavaScript, например используя отдельный файл ресурсов с локализованными строками.

#### <a name="localize-the-add-in-display-name-and-description"></a>Локализация отображаемого имени и описания надстройки

1. В **обозревателе решений** разверните узлы **WorldReadyAddIn** и **WorldReadyAddInManifest**, а затем выберите **WorldReadyAddIn.xml**.

1. В WorldReadyAddInManifest.xml замените [элементы DisplayName] и [Description] следующим блоком кода.

    > [!NOTE]
    > Вы можете заменить локализованные строки на испанском языке, используемые в этом примере для элементов [DisplayName] и [Description], локализованными строками на любом другом языке.

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

1. После изменения отображаемого языка для Office 2013, к примеру, с английского на испанский и последующего запуска надстройки отображаемое имя и описание надстройки локализуются.

#### <a name="lay-out-the-add-in-ui"></a>Разметка пользовательского интерфейса надстройки

1. В **обозревателе решений** Visual Studio выберите элемент **Home.html**.

1. Замените содержимое элемента `<body>` в файле Home.html на приведенный ниже HTML-код и сохраните файл.

    ```html
    <body>
        <!-- Page content -->
        <div id="content-header" class="ms-bgColor-themePrimary ms-font-xl">
            <div class="padding">
                <h1 id="greeting" class="ms-fontColor-white"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div class="ms-font-m">
                    <p id="about"></p>
                </div>
            </div>
        </div>
    </body>
    ```

На приведенном ниже рисунке показаны элемент заголовка (h1) и элемент абзаца (p), в которых будет отображаться локализованный текст после завершения оставшихся действий и запуска надстройки.

*Рис. 1. Пользовательский интерфейс надстройки*

![Пользовательский интерфейс приложения с выделенными разделами.](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a>Добавление файла ресурсов с локализованными строками

Файл ресурсов JavaScript содержит строки, используемые для пользовательского интерфейса надстройки. HTML-код для пользовательского интерфейса примера надстройки содержит элемент `<h1>`, отображающий приветствие, и элемент `<p>`, который знакомит пользователя с надстройкой.

Чтобы включить локализованные строки для заголовка и абзаца, нужно поместить строки в отдельный файл ресурса. Файл ресурса создает объект JavaScript, который содержит отдельный объект Нотация объектов JavaScript (JSON) для каждого набора локализованных строк. Файл ресурса также предоставляет метод для получения соответствующего объекта JSON для определенного региона.

### <a name="add-the-resource-file-to-the-add-in-project"></a>Добавление файла ресурсов в проект надстройки

1. В **обозревателе решений** Visual Studio, щелкните правой кнопкой мыши проект **WorldReadyAddInWeb** и выберите **Добавить** > **Создать элемент**.

1. В диалоговом окне **Добавление нового элемента** выберите параметр **файл JavaScript**.

1. Введите **UIStrings.js** в качестве имени файла и нажмите кнопку **Добавить**.

1. Добавьте в файл UIStrings.js приведенный ниже код и сохраните файл.

    ```js
    /* Store the locale-specific strings */

    const UIStrings = (function ()
    {
        "use strict";

        const UIStrings = {};

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
            let text;

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

Чтобы использовать в надстройке файл ресурсов, вам потребуется добавить для него тег сценария в файл Home.html. При загрузке файла Home.html выполняется файл UIStrings.js, и объект **UIStrings**, используемый для получения строк, становится доступен в коде. Добавьте приведенный ниже HTML-код в тег заголовка для файла Home.html, чтобы сделать объект **UIStrings** доступным в коде.

```html
<!-- Resource file for localized strings: -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

Теперь вы можете использовать объект **UIStrings**, чтобы задать строки для пользовательского интерфейса надстройки.

Если вы хотите изменить локализацию для надстройки в зависимости от того, какой язык используется для отображения в меню и командах в клиентском приложении Office, используйте свойство **Office.context.displayLanguage** , чтобы получить языковой стандарт для этого языка. Например, если язык приложения использует испанский для отображения в меню и командах, свойство **Office.context.displayLanguage** вернет код языка es-ES.

Если вы хотите изменить локализацию для надстройки в зависимости от языка, используемого для редактирования содержимого документа, используйте свойство **Office.context.contentLanguage** , чтобы получить языковой стандарт для этого языка. Например, если язык приложения использует испанский для редактирования содержимого документа, свойство **Office.context.contentLanguage** вернет код языка es-ES.

После того как вы знаете язык, используемый приложением, можно использовать **UIStrings** , чтобы получить набор локализованных строк, соответствующий языку приложения.

Замените код в файле Home.js на следующий код. В коде показано, как можно изменить строки, используемые в элементах пользовательского интерфейса Home.html на основе языка отображения приложения или языка редактирования приложения.

> [!NOTE]
> Чтобы переключаться между локализацией надстройки, основанной на языке редактирования, удалите символы комментария из строки кода `const myLanguage = Office.context.contentLanguage;` и заключите в знаки комментария строку кода `const myLanguage = Office.context.displayLanguage;`.

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {

        $(document).ready(function () {
            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // const myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the Office application.
            const myLanguage = Office.context.displayLanguage;
            let UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Introduction);
        });
    };
})();
```

### <a name="test-your-localized-add-in"></a>Тестирование локализованной надстройки

Чтобы протестировать локализованную надстройку, измените язык, используемый для отображения или редактирования в приложении Office, а затем запустите надстройку.

1. В Word выберите **Файл** > **Параметры** > **Язык**. На рисунке ниже показано диалоговое окно **Параметры Word**, открытое на вкладке языка.

    *Рис. 2. Параметры языка в диалоговом окне "Параметры Word"*

    ![Диалоговое окно "Параметры Word".](../images/office15-app-how-to-localize-fig04.png)

2. В разделе **Выбор языков интерфейса** выберите язык, на котором должны отображаться данные (например, испанский), а затем нажмите стрелку вверх, чтобы переместить испанский язык в начало списка. Кроме того, чтобы изменить язык, используемый для редактирования, в разделе "Выбор языков редактирования" выберите **язык, который** вы хотите использовать для редактирования, например испанский, а затем выберите "Задать по **умолчанию"**.

3. Нажмите кнопку **ОК**, чтобы подтвердить выбор, а затем закройте Word.

4. Нажмите клавишу **F5** в Visual Studio, чтобы запустить пример надстройки, или выберите **Отладка** > **Начать отладку** в строке меню.

5. В Word выберите **Главная** > **Показать область задач**.

После запуска строки в пользовательском интерфейсе надстройки меняются в соответствии с языком, используемым приложением, как показано на следующем рисунке.

*Рис. 3. Пользовательский интерфейс надстройки с локализованным текстом*

![Приложение с локализованным текстом пользовательского интерфейса.](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a>См. также

- [Рекомендации по разработке надстроек Office](../design/add-in-design.md)
- [Идентификаторы языков и значения OptionState Id в Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))

[DefaultLocale]:         /javascript/api/manifest/defaultlocale
[Описание]:           /javascript/api/manifest/description
[DisplayName]:           /javascript/api/manifest/displayname
[IconUrl]:               /javascript/api/manifest/iconurl
[HighResolutionIconUrl]: /javascript/api/manifest/highresolutioniconurl
[Resources]:             /javascript/api/manifest/resources
[SourceLocation]:        /javascript/api/manifest/sourcelocation
[Override]:              /javascript/api/manifest/override
[DesktopSettings]:       /javascript/api/manifest/desktopsettings
[TabletSettings]:        /javascript/api/manifest/tabletsettings
[PhoneSettings]:         /javascript/api/manifest/phonesettings
[displayLanguage]:       /javascript/api/office/office.context#displayLanguage
[contentLanguage]:       /javascript/api/office/office.context#contentLanguage
[RFC 3066]:              https://www.rfc-editor.org/info/rfc3066
