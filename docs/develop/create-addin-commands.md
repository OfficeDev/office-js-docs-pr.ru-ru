---
title: Создание команд надстроек в манифесте для Excel, PowerPoint и Word
description: Используйте VersionOverrides в манифесте, чтобы определить команды надстроек для Excel, PowerPoint и Word. Используйте команды надстроек, чтобы создать элементы пользовательского интерфейса, добавить кнопки или списки, а также для выполнения действий.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 44cd5818879af6788ef58050b5ca475b5f4d3dbd
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889511"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a>Создание команд надстроек в манифесте для Excel, PowerPoint и Word

> [!NOTE]
> В Outlook также поддерживаются команды надстроек. Дополнительные сведения см. в [командах надстройки для Outlook](../outlook/add-in-commands-for-outlook.md)

Используйте **[VersionOverrides](/javascript/api/manifest/versionoverrides)** в манифесте, чтобы определить команды надстроек для Excel, PowerPoint и Word. Команды надстроек позволяют легко настроить пользовательский интерфейс Office по умолчанию, добавив конкретные элементы интерфейса, выполняющие действия. Общие сведения о командах надстроек см. в командах надстроек [для Excel, PowerPoint и Word](../design/add-in-commands.md).

В этой статье описывается изменение манифеста для определения команд надстройки и создание кода для [команд функции](../design/add-in-commands.md#types-of-add-in-commands). На следующей схеме показана иерархия элементов, используемых для задания команд надстройки. Эти элементы подробнее рассматриваются в этой статье.

![Общие сведения об элементах команд надстройки в манифесте. Верхний узел здесь — VersionOverrides с дочерними узлами и ресурсами. В разделе "Узлы" находятся host, а затем DesktopFormFactor. В разделе DesktopFormFactor находятся FunctionFile и ExtensionPoint. В разделе ExtensionPoint находятся customTab или OfficeTab и меню Office. На вкладке CustomTab или Office выберите "Группировать" и "Управление действием". В меню Office выберите пункт "Управление, а затем действие". В разделе Resources (дочерние объекты VersionOverrides) находятся изображения, URL-адреса, shortStrings и LongStrings.](../images/version-overrides.png)

## <a name="step-1-create-the-project"></a>Шаг 1. Создание проекта

Мы рекомендуем создать проект, выполнив одно из кратких руководств, например создание надстройки области [задач Excel](../quickstarts/excel-quickstart-jquery.md). Каждое краткое руководство для Excel, Word и PowerPoint создает проект, который уже содержит команду надстройки (кнопку) для отображения области задач. Прежде чем использовать команды надстроек, убедитесь, что вы считыли команды надстроек для [Excel, Word и PowerPoint](../design/add-in-commands.md) .

## <a name="step-2-create-a-task-pane-add-in"></a>Этап 2. Создание надстройки области задач

Чтобы начать использовать команды надстройки, необходимо сначала создать надстройку области задач, а затем изменить манифест надстройки, как описано в этой статье. С контентными надстройкими нельзя использовать команды надстроек. При обновлении существующего манифеста необходимо добавить соответствующие **пространства имен XML****\<VersionOverrides\>**, а также добавить элемент в манифест, как описано в шаге 3. Добавление [элемента VersionOverrides](#step-3-add-versionoverrides-element).

Ниже приведен пример манифеста надстройки Office 2013. В этом манифесте нет команд надстройки, так как элемент отсутствует **\<VersionOverrides\>** . Office 2013 не поддерживает команды надстройки, **\<VersionOverrides\>** но при добавлении в этот манифест ваша надстройка будет работать как в Office 2013, так и в Office 2016. В Office 2013 надстройка не будет отображать команды надстройки и **\<SourceLocation\>** использует ее значение для запуска надстройки в виде одной надстройки области задач. В Office 2016, **\<VersionOverrides\>** если элемент не включен, в области задач надстройки автоматически открывается URL-адрес, указанный в **\<SourceLocation\>**. Однако если вы включаете **\<VersionOverrides\>** надстройку, она отображает только команды надстройки и изначально не отображает область задач надстройки.
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="https://www.contoso.com/Images/Icon_32.png" />
  <SupportUrl DefaultValue="https://www.contoso.com/contact" />
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
    <AppDomain>AppDomain3</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/Pages/Home.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

 <!-- The VersionOverrides element is inserted at this location in the manifest. -->

</OfficeApp>
```

## <a name="step-3-add-versionoverrides-element"></a>Этап 3. Добавление элемента VersionOverrides

Элемент **\<VersionOverrides\>** является корневым элементом, который содержит определение команды надстройки. **\<VersionOverrides\>** — это дочерний элемент элемента **\<OfficeApp\>** в манифесте. В следующей таблице перечислены атрибуты элемента **\<VersionOverrides\>** .

|Атрибут|Описание|
|:-----|:-----|
|**xmlns** <br/> | Обязательный. Расположение схемы. Необходимое значение — `http://schemas.microsoft.com/office/taskpaneappversionoverrides`. <br/> |
|**xsi:type** <br/> |Обязательный атрибут. Версия схемы. В этой статье описывается версия VersionOverridesV1_0.  <br/> |

В следующей таблице указаны дочерние элементы .**\<VersionOverrides\>**
  
|Элемент|Описание|
|:-----|:-----|
|**\<Description\>** <br/> |Необязательный параметр. Описывает надстройку. Этот дочерний **\<Description\>** элемент переопределяет предыдущий **\<Description\>** элемент в родительской части манифеста. Атрибут **resid** для этого элемента **\<Description\>** имеет идентификатор  **\<String\>** элемента. Элемент **\<String\>** содержит текст для **\<Description\>**. <br/> |
|**\<Requirements\>** <br/> |Необязательный элемент. Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Этот дочерний **\<Requirements\>** элемент переопределяет элемент **\<Requirements\>** в родительской части манифеста. Дополнительные сведения см [. в разделе "Указание приложений Office и требований К API"](../develop/specify-office-hosts-and-api-requirements.md).  <br/> |
|**\<Hosts\>** <br/> |Обязательный. Указывает коллекцию приложений Office. Дочерний **\<Hosts\>** элемент переопределяет элемент **\<Hosts\>** в родительской части манифеста. Необходимо включить атрибут **xsi:type**, для которого задано значение "Книга" или "Документ". <br/> |
|**\<Resources\>** <br/> |Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста. Например, значение элемента **\<Description\>** ссылается на дочерний элемент в **\<Resources\>**. Элемент **\<Resources\>** описан в разделе " [Шаг 7. Добавление элемента Resources](#step-7-add-the-resources-element) " далее в этой статье. <br/> |

В следующем примере показано, как использовать элемент **\<VersionOverrides\>** и его дочерние элементы.

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information about requirement sets -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- add information about form factors -->
      </Host>
      <Host xsi:type="Document">
        <!-- add information about form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information about resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a>Этап 4. Добавление элементов Hosts, Host и DesktopFormFactor

Элемент **\<Hosts\>** содержит один или несколько **\<Host\>** элементов. Элемент **\<Host\>** указывает конкретное приложение Office. Элемент **\<Host\>** содержит дочерние элементы, указывающие команды надстройки, отображаемые после установки надстройки в этом приложении Office. Чтобы отобразить одинаковые команды надстройки в двух или более разных приложениях Office, необходимо дублировать дочерние элементы в каждом из них **\<Host\>**.

Элемент **\<DesktopFormFactor\>** задает параметры для надстройки, которая выполняется в Office в Интернете (в браузере) и Windows.

Ниже приведен пример элементов **\<Hosts\>** и **\<Host\>****\<DesktopFormFactor\>** элементов.

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>

              <!-- information about FunctionFile and ExtensionPoint -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-5-add-the-functionfile-element"></a>Этап 5. Добавление элемента FunctionFile

Элемент **\<FunctionFile\>** указывает файл, содержащий код JavaScript для запуска, когда команда надстройки использует действие **ExecuteFunction** (описание см. в разделе "Элементы управления ["](/javascript/api/manifest/control-button) Кнопка"). Для **\<FunctionFile\>** атрибута **resid** элемента задается HTML-файл, который включает все файлы JavaScript, необходимые командам надстройки. Вы не можете связать непосредственно с файлом JavaScript. Вы можете связать только HTML-файл. Имя файла указывается как элемент **\<Url\>** в элементе **\<Resources\>** .

Ниже приведен пример элемента **\<FunctionFile\>** .
  
```xml
<DesktopFormFactor>
    <FunctionFile resid="residDesktopFuncUrl" />
    <ExtensionPoint xsi:type="PrimaryCommandSurface">
      <!-- information about this extension point -->
    </ExtensionPoint>

    <!-- You can define more than one ExtensionPoint element as needed -->
</DesktopFormFactor>
```

> [!IMPORTANT]
> Убедитесь, что код JavaScript вызывает `Office.initialize`.

Код JavaScript в HTML-файле, на который ссылаются **\<FunctionFile\>** элементы, должен вызываться `Office.initialize`. Элемент **\<FunctionName\>** (см [. элементы управления Button](/javascript/api/manifest/control-button) для описания) использует функции в **\<FunctionFile\>**.

В следующем коде показано, как реализовать функцию, используемую **\<FunctionName\>**.

```js
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Define the function.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("Function command works. Button ID=" + event.source.id,
            function (asyncResult) {
                const error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // Show error message.
                }
                else {
                    // Show success message.
                }
            });

        // Calling event.completed is required. event.completed lets the platform know that processing has completed.
        event.completed();
    }
    
    // You must register the function with the following line.
    Office.actions.associate("writeText", writeText);
</script>
```

> [!IMPORTANT]
> Вызов **event.completed** свидетельствует, что событие успешно обработано. Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь. Первое событие запускается автоматически, тогда как остальные ожидают в очереди. Как только функция вызывает **event.completed**, для нее запускается следующий вызов из очереди. Если объект **event.completed** не реализован, функция не запускается.

## <a name="step-6-add-extensionpoint-elements"></a>Этап 6. Добавление элементов ExtensionPoint

Элемент **\<ExtensionPoint\>** определяет, где должны отображаться команды надстройки в пользовательском интерфейсе Office. Элементы можно определить **\<ExtensionPoint\>** с помощью этих **значений xsi:type** .

- **PrimaryCommandSurface**, которое обозначает ленту в Office.

- **ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в пользовательском интерфейсе Office.

В следующих примерах **\<ExtensionPoint\>** показано, как использовать элемент со значениями атрибутов **PrimaryCommandSurface** и **ContextMenu** , а также дочерними элементами, которые должны использоваться с каждым из них.

> [!IMPORTANT]
> Для элементов, содержащих атрибут идентификатора, необходимо предоставить уникальный идентификатор. Рекомендуем указать название компании с идентификатором. Используйте, например, формат `<CustomTab id="mycompanyname.mygroupname">`.
  
```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso Tab">
  <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
  <!-- <OfficeTab id="TabData"> -->
    <Label resid="residLabel4" />
    <Group id="Group1Id12">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Button1Id1">

        <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

|Элемент|Описание|
|:-----|:-----|
|**\<CustomTab\>** <br/> |Обязательный, если требуется добавить пользовательскую вкладку в ленту (с помощью элемента **PrimaryCommandSurface**). Если вы используете этот **\<CustomTab\>** элемент, его нельзя **\<OfficeTab\>** использовать. Атрибут **id** является обязательным. <br/> |
|**\<OfficeTab\>** <br/> |Требуется, если вы хотите расширить вкладку ленты приложения Office по умолчанию (с помощью **PrimaryCommandSurface**). Если вы используете этот **\<OfficeTab\>** элемент, его нельзя **\<CustomTab\>** использовать. <br/> Дополнительные значения вкладок для использования с атрибутом **id** см. в разделе "Значения табуляции" для вкладок ленты [приложений Office по умолчанию](/javascript/api/manifest/officetab).  <br/> |
|**\<OfficeMenu\>** <br/> | Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: <br/> **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши.<br/> **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь щелкает ячейку электронной таблицы правой кнопкой мыши. <br/> |
|**\<Group\>** <br/> |Группа точек расширения интерфейса пользователя на вкладке. В группе может быть до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов. <br/> |
|**\<Label\>** <br/> |Обязательный. Метка группы. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<ShortStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> |
|**\<Icon\>** <br/> |Обязательный. Определяет значок группы для использования на устройствах с малым форм-фактором или в случаях, когда отображается слишком много кнопок. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<Image\>** . Элемент **\<Image\>** является дочерним элементом элемента **\<Images\>** , который является дочерним элементом элемента **\<Resources\>** . Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64. <br/> |
|**\<Tooltip\>** <br/> |Необязательный параметр. Всплывающая подсказка группы. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<LongStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> |
|**\<Control\>** <br/> |В каждой группе должен быть по крайней мере один элемент управления. Элемент **\<Control\>** может быть кнопкой **или** **меню.** Используйте **меню** , чтобы указать раскрывающийся список элементов управления "Кнопка". В настоящее время поддерживаются только кнопки и меню. [Дополнительные сведения см](/javascript/api/manifest/control-button). в разделах "Элементы управления ["Кнопка](/javascript/api/manifest/control-menu)" и "Меню". <br/>**Примечание:** Чтобы упростить устранение неполадок, **\<Control\>** **\<Resources\>** рекомендуется добавлять элемент и связанные дочерние элементы по одному.          |

### <a name="button-controls"></a>Элементы управления "Кнопка"

Когда пользователь нажимает кнопку, она выполняет одно действие. Она может выполнять функцию JavaScript или отображать область задач. В приведенном ниже примере показано, как определить две кнопки. Первая кнопка выполняет функцию JavaScript без отображения пользовательского интерфейса, а вторая отображает область задач. В элементе **\<Control\>** :

- атрибут **type** является обязательным и должен иметь значение **Button**;

- Атрибут **id** элемента — **\<Control\>** это строка не более 125 символов.

```xml
<!-- Define a control that calls a JavaScript function. -->
<Control xsi:type="Button" id="Button1Id1">
  <Label resid="residLabel" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getData</FunctionName>
  </Action>
</Control>

<!-- Define a control that shows a task pane. -->
<Control xsi:type="Button" id="Button2Id1">
  <Label resid="residLabel2" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon2_32x32" />
    <bt:Image size="32" resid="icon2_32x32" />
    <bt:Image size="80" resid="icon2_32x32" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="residUnitConverterUrl" />
  </Action>
</Control>
```

|Элементы|Описание|
|:-----|:-----|
|**\<Label\>** <br/> |Обязательный. Текст для кнопки. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<ShortStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> |
|**\<Tooltip\>** <br/> |Необязательный параметр. Всплывающая подсказка для кнопки. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<LongStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> |
|**\<Supertip\>** <br/> | Обязательный элемент. Суперподсказка для кнопки, определяемая указанными ниже элементами. <br/> **Title** <br/>  Обязательный. Текст суперподсказки. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<ShortStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> **\<Description\>** <br/>  Обязательный. Описание суперподсказки. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<LongStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> |
|**\<Icon\>** <br/> | Обязательный. Содержит элементы **\<Image\>** для кнопки. Файлы изображений должны быть в формате PNG. <br/> **\<Image\>** <br/>  Определяет изображение для кнопки. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<Image\>** . Элемент **\<Image\>** является дочерним элементом элемента **\<Images\>** , который является дочерним элементом элемента **\<Resources\>** . Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64. <br/> |
|**\<Action\>** <br/> | Обязательный. Указывает действие, которое необходимо выполнить, когда пользователь нажимает кнопку. Для этого атрибута **xsi:type** можно указать следующие значения: <br/> **ExecuteFunction**, который запускает функцию JavaScript, расположенную в файле, на который ссылается .**\<FunctionFile\>** Дочерний **\<FunctionName\>** элемент указывает имя выполняемой функции. <br/> **ShowTaskPane**, отображает область задач надстройки. Дочерний **\<SourceLocation\>** элемент указывает расположение исходного файла отображаемой страницы. Атрибут **resid** должен иметь значение атрибута **идентификатора** **\<Url\>** **\<Urls\>** элемента в элементе **\<Resources\>** элемента. <br/> |

### <a name="menu-controls"></a>Элементы управления "Меню"

Элемент управления **Меню** можно использовать с элементом **PrimaryCommandSurface** или **ContextMenu**. Он определяет следующее:
  
- элемент меню корневого уровня;
- список элементов подменю.

При использовании совместно с элементом **PrimaryCommandSurface**, корневой элемент меню отображается в виде кнопки на ленте. При выборе кнопки отображается подменю в виде раскрывающегося списка. При использовании совместно с элементом **ContextMenu**, элемент меню с подменю вставляется в контекстное меню. В обоих случаях индивидуальные элементы подменю могут выполнять функцию JavaScript или отображать область задач. В настоящее время поддерживается только один уровень подменю.

В приведенном ниже примере показано, как определить элемент меню с двумя элементами подменю. Первый элемент подменю показывает область задач, а второй запускает функцию JavaScript. В элементе **\<Control\>** :

- атрибут **xsi:type** является обязательным и должен иметь значение **Menu**;
- атрибут **id** — это строка длиной до 125 символов.

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

|Элементы|Описание|
|:-----|:-----|
|**\<Label\>** <br/> |Обязательный. Текст корневого элемента меню. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<ShortStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> |
|**\<Tooltip\>** <br/> |Необязательный параметр. Всплывающая подсказка для меню. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<LongStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> |
|**\<SuperTip\>** <br/> | Обязательный элемент. Суперподсказка для меню, определяемая указанными ниже элементами. <br/> **\<Title\>** <br/>  Обязательный. Текст суперподсказки. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<ShortStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> **\<Description\>** <br/>  Обязательный. Описание суперподсказки. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<String\>** . Элемент **\<String\>** является дочерним элементом элемента **\<LongStrings\>** , который является дочерним элементом элемента **\<Resources\>** . <br/> |
|**\<Icon\>** <br/> | Обязательный. Содержит элементы **\<Image\>** меню. Файлы изображений должны быть в формате PNG. <br/> **\<Image\>** <br/>  Изображение для меню. Атрибут **resid** должен иметь значение атрибута **идентификатора** элемента **\<Image\>** . Элемент **\<Image\>** является дочерним элементом элемента **\<Images\>** , который является дочерним элементом элемента **\<Resources\>** . Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения в пикселях: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров в пикселях: 20, 24, 40, 48 и 64. <br/> |
|**\<Items\>** <br/> |Обязательный. Содержит элементы **\<Item\>** для каждого элемента подменю. Каждый **\<Item\>** элемент содержит те же дочерние элементы, что и элементы [управления Button](/javascript/api/manifest/control-button).  <br/> |

## <a name="step-7-add-the-resources-element"></a>Этап 7. Добавление элемента Resources

Элемент **\<Resources\>** содержит ресурсы, используемые различными дочерними элементами **\<VersionOverrides\>** элемента. Ресурсы включают значки, строки и URL-адреса. Элемент манифеста может использовать ресурс, ссылаясь на его **id**. Использование **id** помогает упорядочить манифест, особенно если для разных языковых стандартов используются разные версии ресурса. **id** может содержать до 32 знаков.
  
Ниже приведен пример использования элемента **\<Resources\>** . Каждый ресурс может иметь один или несколько **\<Override\>** дочерних элементов для определения другого ресурса для определенного языкового стандарта.

```xml
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp16-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp32-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp80-icon_default.png" />
    </bt:Image>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
    </bt:Url>
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="residLabel" DefaultValue="GetData">
      <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
    </bt:String>
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="residToolTip" DefaultValue="Get data for your document.">
      <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
    </bt:String>
  </bt:LongStrings>
</Resources>
```

|Ресурс|Описание|
|:-----|:-----|
|**\<Images\>**/ **\<Image\>** <br/> | Предоставляет URL-адрес файла изображения по протоколу HTTPS. Каждое изображение должно быть определено в трех обязательных размерах: <br/>  16×16 <br/>  32×32 <br/>  80×80 <br/>  Кроме того, поддерживаются следующие необязательные размеры: <br/>  20×20 <br/>  24×24 <br/>  40×40 <br/>  48×48 <br/>  64×64 <br/> |
|**\<Urls\>**/ **\<Url\>** <br/> |Предоставляет URL-адрес с префиксом HTTPS. URL-адрес может включать до 2048 символов.  <br/> |
|**\<ShortStrings\>**/ **\<String\>** <br/> |Текст и **\<Label\>** элементы **\<Title\>** . Каждый **\<String\>** из них содержит не более 125 символов. <br/> |
|**\<LongStrings\>**/ **\<String\>** <br/> |Текст и **\<Tooltip\>** элементы **\<Description\>** . Каждый **\<String\>** из них содержит не более 250 символов. <br/> |

> [!NOTE]
> Для всех URL-адресов в элементах и элементах необходимо использовать протокол SSL **\<Image\>** **\<Url\>** .

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a>Значения вкладок для вкладок ленты приложений Office по умолчанию

В Excel и Word вы можете добавить команды надстройки на ленту с помощью стандартных вкладок пользовательского интерфейса Office. В следующей таблице перечислены значения, которые можно использовать для **атрибута идентификатора** элемента **\<OfficeTab\>** . Значения вкладок указываются с учетом регистра.

|Клиентское приложение Office|Значения вкладок|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |

## <a name="see-also"></a>См. также

- [Команды надстроек для Excel, PowerPoint и Word](../design/add-in-commands.md)
- [Пример. Создание надстройки Excel с помощью кнопок команд](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/excel)
- [Пример. Создание надстройки Word с помощью кнопок команд](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/word)
- [Пример. Создание надстройки PowerPoint с помощью кнопок команд](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/powerpoint)
