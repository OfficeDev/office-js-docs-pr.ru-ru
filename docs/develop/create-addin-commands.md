---
title: Создание команд надстройки в манифесте для Excel, PowerPoint и Word
description: Используйте VersionOverrides в манифесте, чтобы определить команды надстройки для Excel, PowerPoint и Word. Используйте команды надстроек для создания элементов пользовательского интерфейса, добавления кнопок или списков, а также для выполнения действий.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 3bcd3c6e07cdb9899601403e68e80e8d609d2e6e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093716"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a><span data-ttu-id="0e948-104">Создание команд надстройки в манифесте для Excel, PowerPoint и Word</span><span class="sxs-lookup"><span data-stu-id="0e948-104">Create add-in commands in your manifest for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="0e948-p102">Используйте **[VersionOverrides](../reference/manifest/versionoverrides.md)** в манифесте, чтобы определить команды надстройки для Excel, PowerPoint и Word. Команды надстроек обеспечивают простой способ настройки пользовательского интерфейса Office по умолчанию с указанными элементами пользовательского интерфейса, которые выполняют действия. С помощью команд надстройки можно:</span><span class="sxs-lookup"><span data-stu-id="0e948-p102">Use **[VersionOverrides](../reference/manifest/versionoverrides.md)** in your manifest to define add-in commands for Excel, PowerPoint, and Word. Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions. You can use add-in commands to:</span></span>

- <span data-ttu-id="0e948-108">Создавать элементы пользовательского интерфейса или точки входа, которые упрощают использование функций надстройки.</span><span class="sxs-lookup"><span data-stu-id="0e948-108">Create UI elements or entry points that make your add-in's functionality easier to use.</span></span>
- <span data-ttu-id="0e948-109">Добавлять кнопки или раскрывающийся список кнопок на ленту.</span><span class="sxs-lookup"><span data-stu-id="0e948-109">Add buttons or a drop-down list of buttons to the ribbon.</span></span>
- <span data-ttu-id="0e948-110">Добавлять отдельные элементы меню, каждый из которых может содержать необязательное подменю, к определенным контекстным меню.</span><span class="sxs-lookup"><span data-stu-id="0e948-110">Add individual menu items — each containing optional submenus — to specific context (shortcut) menus.</span></span>
- <span data-ttu-id="0e948-p103">Выполнять действия при выборе команды надстройки. Варианты действий:</span><span class="sxs-lookup"><span data-stu-id="0e948-p103">Perform actions when your add-in command is chosen. You can:</span></span>
  - <span data-ttu-id="0e948-p104">Показать пользователю одну или несколько надстроек области задач. В надстройке области задач может отображаться код HTML, использующий Office UI Fabric для создания пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="0e948-p104">Show one or more task pane add-ins for users to interact with. Inside your task pane add-in, you can display HTML that uses Office UI Fabric to create a custom UI.</span></span>

     <span data-ttu-id="0e948-115">*или*</span><span class="sxs-lookup"><span data-stu-id="0e948-115">*or*</span></span>

  - <span data-ttu-id="0e948-116">Запустить код JavaScript, который обычно выполняется без отображения пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="0e948-116">Run JavaScript code, which normally runs without displaying any UI.</span></span>

<span data-ttu-id="0e948-p105">В этой статье описывается, как отредактировать манифест, чтобы задать команды надстройки. На следующей схеме показана иерархия элементов, используемых для задания команд надстройки. Эти элементы подробнее рассматриваются в этой статье.</span><span class="sxs-lookup"><span data-stu-id="0e948-p105">This article describes how to edit your manifest to define add-in commands. The following diagram shows the hierarchy of elements used to define add-in commands. These elements are described in more detail in this article.</span></span>

> [!NOTE]
> <span data-ttu-id="0e948-120">В Outlook также поддерживаются команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="0e948-120">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="0e948-121">Дополнительные сведения см в разделе [команды надстроек Outlook](../outlook/add-in-commands-for-outlook.md)</span><span class="sxs-lookup"><span data-stu-id="0e948-121">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md)</span></span>

<span data-ttu-id="0e948-p107">На приведенном ниже изображении представлен обзор элементов команд надстройки в манифесте. ![Обзор элементов команд надстройки в манифесте](../images/version-overrides.png)</span><span class="sxs-lookup"><span data-stu-id="0e948-p107">The following image is an overview of add-in commands elements in the manifest. ![Overview of add-in commands elements in the manifest](../images/version-overrides.png)</span></span>

## <a name="step-1-start-from-a-sample"></a><span data-ttu-id="0e948-124">Этап 1. Ознакомление с примером</span><span class="sxs-lookup"><span data-stu-id="0e948-124">Step 1: Start from a sample</span></span>

<span data-ttu-id="0e948-p108">Настоятельно рекомендуем сначала ознакомиться с одним из примеров, доступных на [странице с примерами команд для надстроек Office](https://github.com/OfficeDev/Office-Add-in-Command-Sample). При необходимости вы можете создать свой манифест, следуя приведенным в руководстве инструкциям. Проверить манифест можно с использованием XSD-файла на сайте с примерами команд для надстроек Office. Прежде чем приступать к использованию команд надстроек, прочтите статью [Команды надстроек для Excel, Word и PowerPoint](../design/add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="0e948-p108">We strongly recommend that you start from one of the samples we provide in  [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Optionally, you can create your own manifest by following the steps in this guide. You can validate your manifest using the XSD file in the Office Add-in Commands Samples site. Ensure that you have read  [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.</span></span>

## <a name="step-2-create-a-task-pane-add-in"></a><span data-ttu-id="0e948-129">Этап 2. Создание надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="0e948-129">Step 2: Create a task pane add-in</span></span>

<span data-ttu-id="0e948-p109">Чтобы приступить к использованию команд надстройки, сначала необходимо создать надстройку области задач, а затем изменить ее манифест, как описано в этой статье. Команды надстроек невозможно использовать с контентными надстройками. Если вы обновляете существующий манифест, добавьте в манифест нужные **пространства имен XML**, а также элемент **VersionOverrides**, как описано в разделе [Шаг 3. Добавление элемента VersionOverrides](#step-3-add-versionoverrides-element).</span><span class="sxs-lookup"><span data-stu-id="0e948-p109">To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article. You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropiate **XML namespaces** as well as add the **VersionOverrides** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).</span></span>

<span data-ttu-id="0e948-p110">Ниже приведен пример манифеста надстройки Office 2013. В этом манифесте нет команд надстройки, так как здесь отсутствует элемент **VersionOverrides**. Office 2013 не поддерживает команды надстройки, но при добавлении элемента **VersionOverrides** в этот манифест надстройка будет работать как в Office 2013, так и в Office 2016. В Office 2013, надстройка не отображает команды и использует значение **SourceLocation** для запуска надстройки в виде единой области задач. В Office 2016, если элемент **VersionOverrides** не включен, для запуска надстройки используется элемент **SourceLocation**. Однако при включении элемента **VersionOverrides** надстройка отображает только команды, но не отображает надстройку в виде единой области задач.</span><span class="sxs-lookup"><span data-stu-id="0e948-p110">The following example shows an Office 2013 add-in's manifest. There are no add-in commands in this manifest because there is no **VersionOverrides** element. Office 2013 doesn't support add-in commands, but by adding **VersionOverrides** to this manifest, your add-in will run in both Office 2013 and Office 2016. In Office 2013, your add-in won't display add-in commands, and uses the value of **SourceLocation** to run your add-in as a single task pane add-in. In Office 2016, if no **VersionOverrides** element is included, **SourceLocation** is used to run your add-in. If you include **VersionOverrides**, however, your add-in displays the add-in commands only, and doesn't display your add-in as a single task pane add-in.</span></span>
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="~remoteAppUrl/Images/Icon_32.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

## <a name="step-3-add-versionoverrides-element"></a><span data-ttu-id="0e948-138">Этап 3. Добавление элемента VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="0e948-138">Step 3: Add VersionOverrides element</span></span>

<span data-ttu-id="0e948-p111">Элемент **VersionOverrides** — это корневой элемент, содержащий определение команды надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента **OfficeApp**. В приведенной ниже таблице перечислены атрибуты элемента **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="0e948-p111">The **VersionOverrides** element is the root element that contains the definition of your add-in command. **VersionOverrides** is a child element of the **OfficeApp** element in the manifest. The following table lists the attributes of the **VersionOverrides** element.</span></span>

|<span data-ttu-id="0e948-142">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="0e948-142">**Attribute**</span></span>|<span data-ttu-id="0e948-143">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0e948-143">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="0e948-144">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="0e948-144">**xmlns**</span></span> <br/> | <span data-ttu-id="0e948-145">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="0e948-145">Required.</span></span> <span data-ttu-id="0e948-146">Расположение схемы. Необходимое значение — `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span><span class="sxs-lookup"><span data-stu-id="0e948-146">The schema location, which must be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span> <br/> |
|<span data-ttu-id="0e948-147">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="0e948-147">**xsi:type**</span></span> <br/> |<span data-ttu-id="0e948-p113">Обязательный атрибут. Версия схемы. В этой статье описывается версия VersionOverridesV1_0.</span><span class="sxs-lookup"><span data-stu-id="0e948-p113">Required. The schema version. The version described in this article is "VersionOverridesV1_0".</span></span>  <br/> |

<span data-ttu-id="0e948-151">В приведенной ниже таблице показаны дочерние элементы **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="0e948-151">The following table identifies the child elements of **VersionOverrides**.</span></span>
  
|<span data-ttu-id="0e948-152">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="0e948-152">**Element**</span></span>|<span data-ttu-id="0e948-153">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0e948-153">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="0e948-154">**Description**</span><span class="sxs-lookup"><span data-stu-id="0e948-154">**Description**</span></span> <br/> |<span data-ttu-id="0e948-p114">Необязательный параметр. Описывает надстройку. Дочерний элемент **Description** переопределяет предыдущий элемент **Description** в родительской части манифеста. Атрибут **resid** для элемента **Description** задан как **id** элемента **String**. Элемент **String** содержит текст для элемента **Description**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p114">Optional. Describes the add-in. This child **Description** element overrides a previous **Description** element in the parent portion of the manifest. The **resid** attribute for this **Description** element is set to the **id** of a **String** element. The **String** element contains the text for **Description**. </span></span><br/> |
|<span data-ttu-id="0e948-160">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="0e948-160">**Requirements**</span></span> <br/> |<span data-ttu-id="0e948-p115">Необязательный параметр. Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Дочерний элемент **Requirements** переопределяет элемент **Requirements** в родительской части манифеста. Дополнительные сведения см. в статье [Указание требований касательно API и узлов Office](../develop/specify-office-hosts-and-api-requirements.md).  </span><span class="sxs-lookup"><span data-stu-id="0e948-p115">Optional. Specifies the minimum requirement set and version of Office.js that the add-in requires. This child **Requirements** element overrides the **Requirements** element in the parent portion of the manifest. For more information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md).  </span></span><br/> |
|<span data-ttu-id="0e948-165">**Hosts**</span><span class="sxs-lookup"><span data-stu-id="0e948-165">**Hosts**</span></span> <br/> |<span data-ttu-id="0e948-p116">Обязательный. Задает набор узлов Office. Дочерний элемент **Hosts** переопределяет элемент **Hosts** в родительской части манифеста. Необходимо включить атрибут **xsi:type**, для которого задано значение "Книга" или "Документ". </span><span class="sxs-lookup"><span data-stu-id="0e948-p116">Required. Specifies a collection of Office hosts. The child **Hosts** element overrides the **Hosts** element in the parent portion of the manifest. You must include a **xsi:type** attribute set to "Workbook" or "Document". </span></span><br/> |
|<span data-ttu-id="0e948-170">**Resources**</span><span class="sxs-lookup"><span data-stu-id="0e948-170">**Resources**</span></span> <br/> |<span data-ttu-id="0e948-p117">Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста. Например, значение элемента **Description** ссылается на дочерний элемент в элементе **Resources**. Элемент **Resources** описан в разделе [Этап 7. Добавление элемента Resources](#step-7-add-the-resources-element) далее в этой статье. </span><span class="sxs-lookup"><span data-stu-id="0e948-p117">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference. For example, the **Description** element's value refers to a child element in **Resources**. The **Resources** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article. </span></span><br/> |

<span data-ttu-id="0e948-174">В приведенном ниже примере показано, как использовать элемент **VersionOverrides** и его дочерние элементы.</span><span class="sxs-lookup"><span data-stu-id="0e948-174">The following example shows how to use the **VersionOverrides** element and its child elements.</span></span>

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

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a><span data-ttu-id="0e948-175">Этап 4. Добавление элементов Hosts, Host и DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="0e948-175">Step 4: Add Hosts, Host, and DesktopFormFactor elements</span></span>

<span data-ttu-id="0e948-p118">Элемент **Hosts** содержит один или несколько элементов **Host**. Элемент **Host** задает конкретный узел Office. Элемент **Host** содержит дочерние элементы, определяющие команды надстройки, которые отображаются после установки надстройки в соответствующем узле Office. Для отображения тех же команд надстройки в нескольких различных узлах Office, необходимо продублировать дочерние элементы в каждом из элементов **Host**.</span><span class="sxs-lookup"><span data-stu-id="0e948-p118">The **Hosts** element contains one or more **Host** elements. A **Host** element specifies a particular Office host. The **Host** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office host. To show the same add-in commands in two or more different Office hosts, you must duplicate the child elements in each **Host**.</span></span>

<span data-ttu-id="0e948-180">Элемент **DesktopFormFactor** задает параметры надстройки, работающей в Office в Интернете (в браузере) и Windows.</span><span class="sxs-lookup"><span data-stu-id="0e948-180">The **DesktopFormFactor** element specifies the settings for an add-in that runs in Office on the web (in a browser) and Windows.</span></span>

<span data-ttu-id="0e948-181">Ниже приведены примеры элементов **Hosts**, **Host** и **DesktopFormFactor**.</span><span class="sxs-lookup"><span data-stu-id="0e948-181">The following is an example of **Hosts**, **Host**, and **DesktopFormFactor** elements.</span></span>

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

## <a name="step-5-add-the-functionfile-element"></a><span data-ttu-id="0e948-182">Этап 5. Добавление элемента FunctionFile</span><span class="sxs-lookup"><span data-stu-id="0e948-182">Step 5: Add the FunctionFile element</span></span>

<span data-ttu-id="0e948-p119">Элемент **FunctionFile** задает файл, который содержит код JavaScript, выполняемый, когда команда надстройки использует действие **ExecuteFunction** (описание см. в разделе [Элементы управления "Кнопка"](../reference/manifest/control.md#button-control)). В атрибуте **resid** элемента **FunctionFile** указан HTML-файл, включающий все файлы JavaScript, необходимые командам надстройки. Ссылаться непосредственно на файл JavaScript невозможно. Вы можете сослаться только на HTML-файл. Имя файла задано в дочернем элементе **Url** элемента **Resources**.</span><span class="sxs-lookup"><span data-stu-id="0e948-p119">The **FunctionFile** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](../reference/manifest/control.md#button-control) for a description). The **FunctionFile** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require. You can't link directly to a JavaScript file. You can only link to an HTML file. The file name is specified as a **Url** element in the **Resources** element.</span></span>

<span data-ttu-id="0e948-188">Ниже приведен пример элемента **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="0e948-188">The following is an example of the **FunctionFile** element.</span></span>
  
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
> <span data-ttu-id="0e948-189">Убедитесь, что код JavaScript вызывает `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="0e948-189">Make sure your JavaScript code calls  `Office.initialize`.</span></span>

<span data-ttu-id="0e948-p120">JavaScript должен вызывать `Office.initialize` в HTML-файле, на который ссылается элемент **FunctionFile**. Элемент **FunctionName** (описание см. в разделе [Элементы управления "Кнопка"](../reference/manifest/control.md#button-control)) использует функции в элементе **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="0e948-p120">The JavaScript in the HTML file referenced by the **FunctionFile** element must call `Office.initialize`. The **FunctionName** element (see [Button controls](../reference/manifest/control.md#button-control) for a description) uses the functions in **FunctionFile**.</span></span>

<span data-ttu-id="0e948-192">Приведенный ниже пример кода показывает, как внедрить функцию, используемую элементом **FunctionName**.</span><span class="sxs-lookup"><span data-stu-id="0e948-192">The following code shows how to implement the function used by **FunctionName**.</span></span>

```js
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Your function must be in the global namespace.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
            function (asyncResult) {
                var error = asyncResult.error;
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
</script>
```

> [!IMPORTANT]
> <span data-ttu-id="0e948-p121">Вызов **event.completed** свидетельствует, что событие успешно обработано. Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь. Первое событие запускается автоматически, тогда как остальные ожидают в очереди. Как только функция вызывает **event.completed**, для нее запускается следующий вызов из очереди. Если объект **event.completed** не реализован, функция не запускается.</span><span class="sxs-lookup"><span data-stu-id="0e948-p121">The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.</span></span>

## <a name="step-6-add-extensionpoint-elements"></a><span data-ttu-id="0e948-198">Этап 6. Добавление элементов ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="0e948-198">Step 6: Add ExtensionPoint elements</span></span>

<span data-ttu-id="0e948-p122">Элемент **ExtensionPoint** определяет, где в пользовательском интерфейсе Office должны появиться команды надстройки. Вы можете определить элементы **ExtensionPoint** по этим значениям **xsi:type**:</span><span class="sxs-lookup"><span data-stu-id="0e948-p122">The **ExtensionPoint** element defines where add-in commands should appear in the Office UI. You can define **ExtensionPoint** elements with these **xsi:type** values:</span></span>

- <span data-ttu-id="0e948-201">**PrimaryCommandSurface**, которое обозначает ленту в Office.</span><span class="sxs-lookup"><span data-stu-id="0e948-201">**PrimaryCommandSurface**, which refers to the ribbon in Office.</span></span>

- <span data-ttu-id="0e948-202">**ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в пользовательском интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="0e948-202">**ContextMenu**, which is the shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="0e948-203">В приведенных ниже примерах показано, как применять элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.</span><span class="sxs-lookup"><span data-stu-id="0e948-203">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0e948-p123">Для элементов, содержащих атрибут идентификатора, необходимо предоставить уникальный идентификатор. Рекомендуем указать название компании с идентификатором. Используйте, например, формат `<CustomTab id="mycompanyname.mygroupname">`.</span><span class="sxs-lookup"><span data-stu-id="0e948-p123">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.</span></span> 
  
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

|<span data-ttu-id="0e948-207">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="0e948-207">**Element**</span></span>|<span data-ttu-id="0e948-208">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0e948-208">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="0e948-209">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="0e948-209">**CustomTab**</span></span> <br/> |<span data-ttu-id="0e948-p124">Обязательный, если требуется добавить пользовательскую вкладку в ленту (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **CustomTab** и **OfficeTab** одновременно. Атрибут **id** является обязательным. </span><span class="sxs-lookup"><span data-stu-id="0e948-p124">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required. </span></span><br/> |
|<span data-ttu-id="0e948-213">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="0e948-213">**OfficeTab**</span></span> <br/> |<span data-ttu-id="0e948-214">Является обязательным, если вы хотите расширить вкладку ленты приложения Office по умолчанию (с помощью **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="0e948-214">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="0e948-215">Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно.</span><span class="sxs-lookup"><span data-stu-id="0e948-215">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <br/> <span data-ttu-id="0e948-216">Для получения дополнительных значений табуляции, используемых с атрибутом **ID** , ознакомьтесь со [значениями вкладок для вкладок приложений Office по умолчанию](../reference/manifest/officetab.md).</span><span class="sxs-lookup"><span data-stu-id="0e948-216">For more tab values to use with the **id** attribute, see [Tab values for default Office app ribbon tabs](../reference/manifest/officetab.md).</span></span>  <br/> |
|<span data-ttu-id="0e948-217">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="0e948-217">**OfficeMenu**</span></span> <br/> | <span data-ttu-id="0e948-p126">Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: </span><span class="sxs-lookup"><span data-stu-id="0e948-p126">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="0e948-p127">**ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="0e948-p127">**ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="0e948-p128">**ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь щелкает ячейку электронной таблицы правой кнопкой мыши. </span><span class="sxs-lookup"><span data-stu-id="0e948-p128">**ContextMenuCell** for Excel. Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet. </span></span><br/> |
|<span data-ttu-id="0e948-224">**Group**</span><span class="sxs-lookup"><span data-stu-id="0e948-224">**Group**</span></span> <br/> |<span data-ttu-id="0e948-p129">Группа точек расширения интерфейса пользователя на вкладке. В группе может быть до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов. </span><span class="sxs-lookup"><span data-stu-id="0e948-p129">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="0e948-228">**Label**</span><span class="sxs-lookup"><span data-stu-id="0e948-228">**Label**</span></span> <br/> |<span data-ttu-id="0e948-p130">Обязательный. Метка группы. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p130">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="0e948-233">**Icon**</span><span class="sxs-lookup"><span data-stu-id="0e948-233">**Icon**</span></span> <br/> |<span data-ttu-id="0e948-p131">Обязательный. Определяет значок группы для использования на устройствах с малым форм-фактором или в случаях, когда отображается слишком много кнопок. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **Image**. **Image** — это дочерний элемент **Images**, который в свою очередь является дочерним для элемента **Resources**. Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64. </span><span class="sxs-lookup"><span data-stu-id="0e948-p131">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="0e948-241">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="0e948-241">**Tooltip**</span></span> <br/> |<span data-ttu-id="0e948-p132">Необязательный параметр. Всплывающая подсказка группы. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p132">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="0e948-246">**Control**</span><span class="sxs-lookup"><span data-stu-id="0e948-246">**Control**</span></span> <br/> |<span data-ttu-id="0e948-p133">Для каждой группы требуется хотя бы один элемент управления. Элемент **Control** может иметь значение **Button** или **Menu**. Укажите **Menu**, чтобы задать раскрывающийся список элементов управления "Кнопка". В настоящий момент поддерживаются только кнопки и меню. Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](../reference/manifest/control.md#button-control) и [Элементы управления "Меню"](../reference/manifest/control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="0e948-p133">Each group requires at least one control. A **Control** element can be either a **Button** or a **Menu**. Use **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported. See the  [Button controls](../reference/manifest/control.md#button-control) and [Menu controls](../reference/manifest/control.md#menu-dropdown-button-controls) sections for more information. </span></span><br/><span data-ttu-id="0e948-252">**Примечание.** Чтобы упростить устранение неполадок, рекомендуем добавлять элемент **Control** и соответствующие дочерние элементы **Resources** по одному.</span><span class="sxs-lookup"><span data-stu-id="0e948-252">**Note:** To make troubleshooting easier, we recommend that you add a **Control** element and the related **Resources** child elements one at a time.</span></span>          |

### <a name="button-controls"></a><span data-ttu-id="0e948-253">Элементы управления "Кнопка"</span><span class="sxs-lookup"><span data-stu-id="0e948-253">Button controls</span></span>

<span data-ttu-id="0e948-p134">Когда пользователь нажимает кнопку, она выполняет одно действие. Она может выполнять функцию JavaScript или отображать область задач. В приведенном ниже примере показано, как определить две кнопки. Первая кнопка выполняет функцию JavaScript без отображения пользовательского интерфейса, а вторая отображает область задач. В элементе **Control**:</span><span class="sxs-lookup"><span data-stu-id="0e948-p134">A button performs a single action when the user selects it. It can either execute a JavaScript function or show a task pane. The following example shows how to define two buttons. The first button runs a JavaScript function without showing a UI, and the second button shows a task pane. In the **Control** element:</span></span>

- <span data-ttu-id="0e948-259">атрибут **type** является обязательным и должен иметь значение **Button**;</span><span class="sxs-lookup"><span data-stu-id="0e948-259">The **type** attribute is required, and must be set to **Button**.</span></span>

- <span data-ttu-id="0e948-260">атрибут **id** элемента **Control** — это строка длиной до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="0e948-260">The **id** attribute of the **Control** element is a string with a maximum of 125 characters.</span></span>

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

|<span data-ttu-id="0e948-261">**Элементы**</span><span class="sxs-lookup"><span data-stu-id="0e948-261">**Elements**</span></span>|<span data-ttu-id="0e948-262">**Description**</span><span class="sxs-lookup"><span data-stu-id="0e948-262">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="0e948-263">**Label**</span><span class="sxs-lookup"><span data-stu-id="0e948-263">**Label**</span></span> <br/> |<span data-ttu-id="0e948-p135">Обязательный. Текст для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p135">Required. The text for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="0e948-268">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="0e948-268">**Tooltip**</span></span> <br/> |<span data-ttu-id="0e948-p136">Необязательный параметр. Всплывающая подсказка для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p136">Optional. The tooltip for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="0e948-273">**Supertip**</span><span class="sxs-lookup"><span data-stu-id="0e948-273">**Supertip**</span></span> <br/> | <span data-ttu-id="0e948-p137">Обязательный элемент. Суперподсказка для кнопки, определяемая указанными ниже элементами. </span><span class="sxs-lookup"><span data-stu-id="0e948-p137">Required. The supertip for this button, which is defined by the following: </span></span><br/> <span data-ttu-id="0e948-276">**Title**</span><span class="sxs-lookup"><span data-stu-id="0e948-276">**Title**</span></span> <br/>  <span data-ttu-id="0e948-p138">Обязательный. Текст суперподсказки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p138">Required. The text for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="0e948-281">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0e948-281">**Description**</span></span> <br/>  <span data-ttu-id="0e948-p139">Обязательный. Описание суперподсказки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p139">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="0e948-286">**Icon**</span><span class="sxs-lookup"><span data-stu-id="0e948-286">**Icon**</span></span> <br/> | <span data-ttu-id="0e948-p140">Обязательный. Содержит элементы **Image** для кнопки. Файлы изображений должны быть в формате PNG. </span><span class="sxs-lookup"><span data-stu-id="0e948-p140">Required. Contains the **Image** elements for the button. Image files must be .png format. </span></span><br/> <span data-ttu-id="0e948-290">**Image**</span><span class="sxs-lookup"><span data-stu-id="0e948-290">**Image**</span></span> <br/>  <span data-ttu-id="0e948-p141">Определяет изображение для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **Image**. **Image** — это дочерний элемент **Images**, который в свою очередь является дочерним для элемента **Resources**. Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64. </span><span class="sxs-lookup"><span data-stu-id="0e948-p141">Defines an image to display on the button. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="0e948-297">**Действие**</span><span class="sxs-lookup"><span data-stu-id="0e948-297">**Action**</span></span> <br/> | <span data-ttu-id="0e948-p142">Обязательный. Указывает действие, которое необходимо выполнить, когда пользователь нажимает кнопку. Для этого атрибута **xsi:type** можно указать следующие значения: </span><span class="sxs-lookup"><span data-stu-id="0e948-p142">Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: </span></span><br/> <span data-ttu-id="0e948-p143">**ExecuteFunction.** Вызывает функцию JavaScript, расположенную в файле, на который ссылается элемент **FunctionFile**. **ExecuteFunction** не отображает пользовательский интерфейс. Дочерний элемент **FunctionName** задает имя выполняемой функции.</span><span class="sxs-lookup"><span data-stu-id="0e948-p143">**ExecuteFunction**, which runs a JavaScript function located in the file referenced by **FunctionFile**. **ExecuteFunction** does not display a UI. The **FunctionName** child element specifies the name of the function to execute. </span></span><br/> <span data-ttu-id="0e948-p144">**ShowTaskPane.** Отображает надстройку области задач. Дочерний элемент **SourceLocation** задает расположение исходного файла отображаемой надстройки области задач. Для атрибута **resid** должно быть задано значение атрибута **id** элемента **Url** в элементе **Urls**, включенном в элемент **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p144">**ShowTaskPane**, which shows a task pane add-in. The **SourceLocation** child element specifies the source file location of the task pane add-in to display. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element. </span></span><br/> |

### <a name="menu-controls"></a><span data-ttu-id="0e948-307">Элементы управления "Меню"</span><span class="sxs-lookup"><span data-stu-id="0e948-307">Menu controls</span></span>

<span data-ttu-id="0e948-308">Элемент управления **Меню** можно использовать с элементом **PrimaryCommandSurface** или **ContextMenu**. Он определяет следующее:</span><span class="sxs-lookup"><span data-stu-id="0e948-308">A **Menu** control can be used with either **PrimaryCommandSurface** or **ContextMenu**, and defines:</span></span>
  
- <span data-ttu-id="0e948-309">элемент меню корневого уровня;</span><span class="sxs-lookup"><span data-stu-id="0e948-309">A root-level menu item.</span></span>
- <span data-ttu-id="0e948-310">список элементов подменю.</span><span class="sxs-lookup"><span data-stu-id="0e948-310">A list of submenu items.</span></span>

<span data-ttu-id="0e948-p145">При использовании совместно с элементом **PrimaryCommandSurface**, корневой элемент меню отображается в виде кнопки на ленте. При выборе кнопки отображается подменю в виде раскрывающегося списка. При использовании совместно с элементом **ContextMenu**, элемент меню с подменю вставляется в контекстное меню. В обоих случаях индивидуальные элементы подменю могут выполнять функцию JavaScript или отображать область задач. В настоящее время поддерживается только один уровень подменю.</span><span class="sxs-lookup"><span data-stu-id="0e948-p145">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="0e948-p146">В приведенном ниже примере показано, как определить элемент меню с двумя элементами подменю. Первый элемент подменю показывает область задач, а второй запускает функцию JavaScript. В элементе **Control**:</span><span class="sxs-lookup"><span data-stu-id="0e948-p146">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function. In the **Control** element:</span></span>

- <span data-ttu-id="0e948-319">атрибут **xsi:type** является обязательным и должен иметь значение **Menu**;</span><span class="sxs-lookup"><span data-stu-id="0e948-319">The **xsi:type** attribute is required, and must be set to **Menu**.</span></span>
- <span data-ttu-id="0e948-320">атрибут **id** — это строка длиной до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="0e948-320">The **id** attribute is a string with a maximum of 125 characters.</span></span>

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

|<span data-ttu-id="0e948-321">**Элементы**</span><span class="sxs-lookup"><span data-stu-id="0e948-321">**Elements**</span></span>|<span data-ttu-id="0e948-322">**Description**</span><span class="sxs-lookup"><span data-stu-id="0e948-322">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="0e948-323">**Label**</span><span class="sxs-lookup"><span data-stu-id="0e948-323">**Label**</span></span> <br/> |<span data-ttu-id="0e948-p147">Обязательный. Текст корневого элемента меню. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p147">Required. The text of the root menu item. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="0e948-328">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="0e948-328">**Tooltip**</span></span> <br/> |<span data-ttu-id="0e948-p148">Необязательный параметр. Всплывающая подсказка для меню. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p148">Optional. The tooltip for the menu. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="0e948-333">**SuperTip**</span><span class="sxs-lookup"><span data-stu-id="0e948-333">**SuperTip**</span></span> <br/> | <span data-ttu-id="0e948-p149">Обязательный элемент. Суперподсказка для меню, определяемая указанными ниже элементами. </span><span class="sxs-lookup"><span data-stu-id="0e948-p149">Required. The supertip for the menu, which is defined by the following: </span></span><br/> <span data-ttu-id="0e948-336">**Title**</span><span class="sxs-lookup"><span data-stu-id="0e948-336">**Title**</span></span> <br/>  <span data-ttu-id="0e948-p150">Обязательный. Текст суперподсказки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p150">Required. The text of the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="0e948-341">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0e948-341">**Description**</span></span> <br/>  <span data-ttu-id="0e948-p151">Обязательный. Описание суперподсказки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="0e948-p151">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="0e948-346">**Icon**</span><span class="sxs-lookup"><span data-stu-id="0e948-346">**Icon**</span></span> <br/> | <span data-ttu-id="0e948-p152">Обязательный. Содержит элементы **Image** для меню. Файлы изображений должны быть в формате PNG. </span><span class="sxs-lookup"><span data-stu-id="0e948-p152">Required. Contains the **Image** elements for the menu. Image files must be .png format. </span></span><br/> <span data-ttu-id="0e948-350">**Image**</span><span class="sxs-lookup"><span data-stu-id="0e948-350">**Image**</span></span> <br/>  <span data-ttu-id="0e948-p153">Изображение для меню. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **Image**. **Image** — это дочерний элемент **Images**, который в свою очередь является дочерним для элемента **Resources**. Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения в пикселях: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров в пикселях: 20, 24, 40, 48 и 64. </span><span class="sxs-lookup"><span data-stu-id="0e948-p153">An image for the menu. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes, in pixels, are required: 16, 32, and 80. Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="0e948-357">**Items**</span><span class="sxs-lookup"><span data-stu-id="0e948-357">**Items**</span></span> <br/> |<span data-ttu-id="0e948-p154">Обязательный. Содержит элементы **Item** для каждого элемента подменю. Каждый элемент **Item** содержит те же дочерние элементы, что и [Элементы управления ''Кнопка''](../reference/manifest/control.md#button-control).  </span><span class="sxs-lookup"><span data-stu-id="0e948-p154">Required. Contains the **Item** elements for each submenu item. Each **Item** element contains the same child elements as [Button controls](../reference/manifest/control.md#button-control).  </span></span><br/> |

## <a name="step-7-add-the-resources-element"></a><span data-ttu-id="0e948-361">Этап 7. Добавление элемента Resources</span><span class="sxs-lookup"><span data-stu-id="0e948-361">Step 7: Add the Resources element</span></span>

<span data-ttu-id="0e948-p155">Элемент **Resources** содержит ресурсы, используемые различными дочерними элементами элемента **VersionOverrides**. Ресурсы включают значки, строки и URL-адреса. Элемент манифеста может использовать ресурс, ссылаясь на его **id**. Использование **id** помогает упорядочить манифест, особенно если для разных языковых стандартов используются разные версии ресурса. **id** может содержать до 32 знаков.</span><span class="sxs-lookup"><span data-stu-id="0e948-p155">The **Resources** element contains resources used by the different child elements of the **VersionOverrides** element. Resources include icons, strings, and URLs. An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.</span></span>
  
<span data-ttu-id="0e948-p156">Ниже приведен пример использования элемента **Resources**. Каждый ресурс может иметь один или несколько дочерних элементов **Override**, позволяющих указать другой ресурс для определенного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="0e948-p156">The following shows an example of how to use the **Resources** element. Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

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

|<span data-ttu-id="0e948-369">**Ресурс**</span><span class="sxs-lookup"><span data-stu-id="0e948-369">**Resource**</span></span>|<span data-ttu-id="0e948-370">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0e948-370">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="0e948-371">**Images**/ **Image**</span><span class="sxs-lookup"><span data-stu-id="0e948-371">**Images**/ **Image**</span></span> <br/> | <span data-ttu-id="0e948-p157">Предоставляет URL-адрес файла изображения по протоколу HTTPS. Каждое изображение должно быть определено в трех обязательных размерах:</span><span class="sxs-lookup"><span data-stu-id="0e948-p157">Provides the HTTPS URL to an image file. Each image must define the three required image sizes:</span></span> <br/>  <span data-ttu-id="0e948-374">16×16</span><span class="sxs-lookup"><span data-stu-id="0e948-374">16×16</span></span> <br/>  <span data-ttu-id="0e948-375">32×32</span><span class="sxs-lookup"><span data-stu-id="0e948-375">32×32</span></span> <br/>  <span data-ttu-id="0e948-376">80×80</span><span class="sxs-lookup"><span data-stu-id="0e948-376">80×80</span></span> <br/>  <span data-ttu-id="0e948-377">Кроме того, поддерживаются следующие необязательные размеры:</span><span class="sxs-lookup"><span data-stu-id="0e948-377">The following image sizes are also supported, but not required:</span></span> <br/>  <span data-ttu-id="0e948-378">20×20</span><span class="sxs-lookup"><span data-stu-id="0e948-378">20×20</span></span> <br/>  <span data-ttu-id="0e948-379">24×24</span><span class="sxs-lookup"><span data-stu-id="0e948-379">24×24</span></span> <br/>  <span data-ttu-id="0e948-380">40×40</span><span class="sxs-lookup"><span data-stu-id="0e948-380">40×40</span></span> <br/>  <span data-ttu-id="0e948-381">48×48</span><span class="sxs-lookup"><span data-stu-id="0e948-381">48×48</span></span> <br/>  <span data-ttu-id="0e948-382">64×64</span><span class="sxs-lookup"><span data-stu-id="0e948-382">64×64</span></span> <br/> |
|<span data-ttu-id="0e948-383">**Urls**/ **Url**</span><span class="sxs-lookup"><span data-stu-id="0e948-383">**Urls**/ **Url**</span></span> <br/> |<span data-ttu-id="0e948-p158">Предоставляет URL-адрес с префиксом HTTPS. URL-адрес может включать до 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="0e948-p158">Provides an HTTPS URL location. A URL can be a maximum of 2048 characters.</span></span>  <br/> |
|<span data-ttu-id="0e948-386">**ShortStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="0e948-386">**ShortStrings**/ **String**</span></span> <br/> |<span data-ttu-id="0e948-p159">Текст для элементов **Label** и **Title**. Каждая **строка** содержит не более 125 символов. </span><span class="sxs-lookup"><span data-stu-id="0e948-p159">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="0e948-389">**LongStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="0e948-389">**LongStrings**/ **String**</span></span> <br/> |<span data-ttu-id="0e948-p160">Текст для элементов **Tooltip** и **Description**. Каждый элемент **String** содержит не более 250 символов.</span><span class="sxs-lookup"><span data-stu-id="0e948-p160">The text for **Tooltip** and **Description** elements. Each **String** contains a maximum of 250 characters. </span></span><br/> |

> [!NOTE]
> <span data-ttu-id="0e948-392">Для всех URL-адресов в элементах **Image** и **Url** необходимо использовать протокол SSL.</span><span class="sxs-lookup"><span data-stu-id="0e948-392">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a><span data-ttu-id="0e948-393">Значения табуляции на вкладках ленты приложения Office по умолчанию</span><span class="sxs-lookup"><span data-stu-id="0e948-393">Tab values for default Office app ribbon tabs</span></span>

<span data-ttu-id="0e948-p161">В Excel и Word вы можете добавить команды надстройки на ленту с помощью стандартных вкладок пользовательского интерфейса Office. В приведенной ниже таблице перечислены значения, которые можно использовать для атрибута **id** элемента **OfficeTab**. Значения вкладок указываются с учетом регистра.</span><span class="sxs-lookup"><span data-stu-id="0e948-p161">In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs. The following table lists the values that you can use for the **id** attribute of the **OfficeTab** element. The tab values are case sensitive.</span></span>

|<span data-ttu-id="0e948-397">**Ведущее приложение Office**</span><span class="sxs-lookup"><span data-stu-id="0e948-397">**Office host application**</span></span>|<span data-ttu-id="0e948-398">**Значения вкладок**</span><span class="sxs-lookup"><span data-stu-id="0e948-398">**Tab values**</span></span>|
|:-----|:-----|
|<span data-ttu-id="0e948-399">Excel</span><span class="sxs-lookup"><span data-stu-id="0e948-399">Excel</span></span>  <br/> |<span data-ttu-id="0e948-400">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span><span class="sxs-lookup"><span data-stu-id="0e948-400">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span></span> <br/> |
|<span data-ttu-id="0e948-401">Word</span><span class="sxs-lookup"><span data-stu-id="0e948-401">Word</span></span>  <br/> |<span data-ttu-id="0e948-402">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span><span class="sxs-lookup"><span data-stu-id="0e948-402">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span></span> <br/> |
|<span data-ttu-id="0e948-403">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0e948-403">PowerPoint</span></span>  <br/> |<span data-ttu-id="0e948-404">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span><span class="sxs-lookup"><span data-stu-id="0e948-404">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span></span>          <br/> |

## <a name="see-also"></a><span data-ttu-id="0e948-405">См. также</span><span class="sxs-lookup"><span data-stu-id="0e948-405">See also</span></span>

- [<span data-ttu-id="0e948-406">Команды надстроек для Excel, PowerPoint и Word</span><span class="sxs-lookup"><span data-stu-id="0e948-406">Add-in commands for Excel, PowerPoint, and Word</span></span>](../design/add-in-commands.md)
