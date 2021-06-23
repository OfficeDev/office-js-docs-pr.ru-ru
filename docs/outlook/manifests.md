---
title: Манифесты надстройки Outlook
description: В манифесте описывается, как выполняется интеграция надстройки Outlook с клиентами Outlook, включая пример.
ms.date: 05/27/2020
localization_priority: Priority
ms.openlocfilehash: f113a5d8f92ee80ed635283e9e5544bd4b9ce7cd
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076772"
---
# <a name="outlook-add-in-manifests"></a><span data-ttu-id="42137-103">Манифесты надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="42137-103">Outlook add-in manifests</span></span>

<span data-ttu-id="42137-p101">Надстройка Outlook состоит из двух компонентов: XML-манифеста надстройки и веб-страницы с поддержкой библиотеки JavaScript для надстроек Office (office.js). В манифесте описывается интеграция надстройки с клиентами Outlook. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="42137-p101">An Outlook add-in consists of two components: the XML add-in manifest and a web page supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients. The following is an example.</span></span>

 > [!NOTE]
 > <span data-ttu-id="42137-p102">Все значения URL-адресов в следующем примере начинаются со строки "https://appdemo.contoso.com". Это значение — заполнитель. В фактическом допустимом манифесте эти значения будут содержать действительные URL-адреса с префиксом HTTPS.</span><span class="sxs-lookup"><span data-stu-id="42137-p102">All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="https://appdemo.contoso.com/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="https://appdemo.contoso.com/images/blue-128.png" />
  <SupportUrl DefaultValue="https://appdemo.contoso.com"/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read task pane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="functionFile" />

          <!-- Message read form -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadDemoGroup">
                <Label resid="groupLabel" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Supertip>
                    <Title resid="funcReadSuperTipTitle" />
                    <Description resid="funcReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getSubject</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="msgReadMenuButton">
                  <Label resid="menuReadButtonLabel" />
                  <Supertip>
                    <Title resid="menuReadSuperTipTitle" />
                    <Description resid="menuReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="msgReadMenuItem1">
                      <Label resid="menuItem1ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem1ReadLabel" />
                        <Description resid="menuItem1ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemClass</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem2ReadLabel" />
                        <Description resid="menuItem2ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getDateTimeCreated</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem3ReadLabel" />
                        <Description resid="menuItem3ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemID</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Supertip>
                    <Title resid="paneReadSuperTipTitle" />
                    <Description resid="paneReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="readTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <!-- Blue icon -->
        <bt:Image id="blue-icon-16" DefaultValue="https://appdemo.contoso.com/images/blue-16.png" />
        <bt:Image id="blue-icon-32" DefaultValue="https://appdemo.contoso.com/images/blue-32.png" />
        <bt:Image id="blue-icon-80" DefaultValue="https://appdemo.contoso.com/images/blue-80.png" />
        <!-- Red icon -->
        <bt:Image id="red-icon-16" DefaultValue="https://appdemo.contoso.com/images/red-16.png" />
        <bt:Image id="red-icon-32" DefaultValue="https://appdemo.contoso.com/images/red-32.png" />
        <bt:Image id="red-icon-80" DefaultValue="https://appdemo.contoso.com/images/red-80.png" />
        <!-- Green icon -->
        <bt:Image id="green-icon-16" DefaultValue="https://appdemo.contoso.com/images/green-16.png" />
        <bt:Image id="green-icon-32" DefaultValue="https://appdemo.contoso.com/images/green-32.png" />
        <bt:Image id="green-icon-80" DefaultValue="https://appdemo.contoso.com/images/green-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="https://appdemo.contoso.com/FunctionFile/Functions.html" />
        <bt:Url id="readTaskPaneUrl" DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
        <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
        <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

        <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
        <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

        <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
        <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
        <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
        <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

        <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
        <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
        <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

## <a name="schema-versions"></a><span data-ttu-id="42137-110">Версии схемы</span><span class="sxs-lookup"><span data-stu-id="42137-110">Schema versions</span></span>

<span data-ttu-id="42137-p103">Не все клиенты Outlook поддерживают новейшие функции, а у некоторых пользователей Outlook установлена более ранняя версия Outlook. С помощью версий схемы разработчики могут создавать надстройки с обратной совместимостью, которые используют новые функции, если они доступны, но работают и в более ранних версиях.</span><span class="sxs-lookup"><span data-stu-id="42137-p103">Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.</span></span>

<span data-ttu-id="42137-p104">Наглядный пример — элемент манифеста **VersionOverrides**. Все элементы, определенные в **VersionOverrides**, заменяют соответствующий элемент в другой части манифеста. Это означает, что по мере возможности Outlook будет использовать содержимое раздела **VersionOverrides** для установки параметров надстройки. Тем не менее, если версия Outlook не поддерживает определенную версию **VersionOverrides**, Outlook пропустит ее и будет использовать остальные сведения из манифеста.</span><span class="sxs-lookup"><span data-stu-id="42137-p104">The **VersionOverrides** element in the manifest is an example of this. All elements defined inside **VersionOverrides** will override the same element in the other part of the manifest. This means that, whenever possible, Outlook will use what is in the **VersionOverrides** section to set up the add-in. However, if the version of Outlook doesn't support a certain version of **VersionOverrides**, Outlook will ignore it and depend on the information in the rest of the manifest.</span></span> 

<span data-ttu-id="42137-117">Этот подход означает, что разработчикам не требуется создавать несколько отдельных манифестов — все параметры можно задать в одном файле.</span><span class="sxs-lookup"><span data-stu-id="42137-117">This approach means that developers don't have to create multiple individual manifests, but rather keep everything defined in one file.</span></span>

<span data-ttu-id="42137-118">В настоящий момент доступны следующие версии схемы:</span><span class="sxs-lookup"><span data-stu-id="42137-118">The current versions of the schema are:</span></span>


|<span data-ttu-id="42137-119">Версия</span><span class="sxs-lookup"><span data-stu-id="42137-119">Version</span></span>|<span data-ttu-id="42137-120">Описание</span><span class="sxs-lookup"><span data-stu-id="42137-120">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="42137-121">v1.0</span><span class="sxs-lookup"><span data-stu-id="42137-121">v1.0</span></span>|<span data-ttu-id="42137-p105">Поддерживает версию 1.0 API JavaScript для Office. Для надстроек Outlook поддерживается форма чтения.</span><span class="sxs-lookup"><span data-stu-id="42137-p105">Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form.</span></span> |
|<span data-ttu-id="42137-124">1.1</span><span class="sxs-lookup"><span data-stu-id="42137-124">v1.1</span></span>|<span data-ttu-id="42137-p106">Поддерживает версии 1.1 API JavaScript для Office и **VersionOverrides**. Для надстроек Outlook поддерживается форма создания.</span><span class="sxs-lookup"><span data-stu-id="42137-p106">Supports version 1.1 of the Office JavaScript API and **VersionOverrides**. For Outlook add-ins, this adds support for compose form.</span></span>|
|<span data-ttu-id="42137-127">**VersionOverrides** 1.0</span><span class="sxs-lookup"><span data-stu-id="42137-127">**VersionOverrides** 1.0</span></span>|<span data-ttu-id="42137-p107">Поддерживает более поздние версии API JavaScript для Office. Поддерживаются команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="42137-p107">Supports later versions of the Office JavaScript API. This supports add-in commands.</span></span>|
|<span data-ttu-id="42137-130">**VersionOverrides** 1.1</span><span class="sxs-lookup"><span data-stu-id="42137-130">**VersionOverrides** 1.1</span></span>|<span data-ttu-id="42137-p108">Поддерживает более поздние версии API JavaScript для Office. Поддерживает команды надстроек и добавляет поддержку новых функций, таких как [закрепляемые области задач](pinnable-taskpane.md) и мобильные надстройки.</span><span class="sxs-lookup"><span data-stu-id="42137-p108">Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.</span></span>|

<span data-ttu-id="42137-p109">В этой статье рассматриваются требования для манифеста версии 1.1. Даже если в манифесте вашей надстройки используется элемент **VersionOverrides**, все равно важно включить элементы манифеста версии 1.1, чтобы надстройка работала со старыми клиентами, которые не поддерживают **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="42137-p109">This article will cover the requirements for a v1.1 manifest. Even if your add-in manifest uses the **VersionOverrides** element, it is still important to include the v1.1 manifest elements to allow your add-in to work with older clients that do not support **VersionOverrides**.</span></span>

> [!NOTE]
> <span data-ttu-id="42137-p110">Outlook использует схему для проверки манифестов. Поэтому элементы манифеста должны располагаться в определенном порядке. Если порядок не соблюден, при загрузке неопубликованной надстройки могут возникать ошибки. Вы можете скачать [определение схемы XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), которое позволит создать манифест с правильным расположением элементов.</span><span class="sxs-lookup"><span data-stu-id="42137-p110">Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.</span></span>

## <a name="root-element"></a><span data-ttu-id="42137-139">Корневой элемент</span><span class="sxs-lookup"><span data-stu-id="42137-139">Root element</span></span>

<span data-ttu-id="42137-p111">Корневой элемент манифеста надстройки Outlook — **OfficeApp**. Этот элемент также объявляет пространство имен по умолчанию, версию схемы и тип надстройки. Поместите все остальные элементы манифеста между его открывающим и закрывающим тегами. Ниже приводится пример корневого элемента.</span><span class="sxs-lookup"><span data-stu-id="42137-p111">The root element for the Outlook add-in manifest is **OfficeApp**. This element also declares the default namespace, schema version and the type of add-in. Place all other elements in the manifest within its open and close tags. The following is an example of the root element:</span></span>


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest -->

</OfficeApp>
```

## <a name="version"></a><span data-ttu-id="42137-144">Version</span><span class="sxs-lookup"><span data-stu-id="42137-144">Version</span></span>

<span data-ttu-id="42137-p112">Это версия конкретной надстройки. Когда разработчик обновляет какой-либо элемент манифеста, версию также необходимо увеличить. Таким образом, при установке нового манифеста заменяется имеющийся, а пользователю становятся доступны новые функции. Если эта надстройка была отправлена в магазин, манифест потребуется заново отправить и проверить. Спустя несколько часов (после утверждения обновленного манифеста) пользователи надстройки автоматически получат его.</span><span class="sxs-lookup"><span data-stu-id="42137-p112">This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.</span></span>

<span data-ttu-id="42137-p113">При изменении разрешений, запрашиваемых надстройкой, пользователям предлагается выполнить обновление и повторно согласиться на предоставление надстройке разрешений. Если администратор установил эту надстройку для всей организации, он должен будет дать свое согласие первым. До этого пользователям будут доступны только старые функции.</span><span class="sxs-lookup"><span data-stu-id="42137-p113">If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.</span></span>

## <a name="versionoverrides"></a><span data-ttu-id="42137-153">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="42137-153">VersionOverrides</span></span>

<span data-ttu-id="42137-154">Элемент **VersionOverrides** — это расположение данных о [командах надстройки](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="42137-154">The **VersionOverrides** element is the location of information for [add-in commands](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="42137-155">Кроме того, в этом элементе определяется поддержка [мобильных надстроек](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="42137-155">This element is also where add-ins define support for [mobile add-ins](add-mobile-support.md).</span></span>

<span data-ttu-id="42137-156">Описание этого элемента см. в статье [Создание команд надстроек в манифесте для Excel, PowerPoint и Word](../develop/create-addin-commands.md).</span><span class="sxs-lookup"><span data-stu-id="42137-156">For a discussion on this element, see [Create add-in commands in your manifest for Excel, PowerPoint, and Word](../develop/create-addin-commands.md).</span></span>

## <a name="localization"></a><span data-ttu-id="42137-157">Локализация</span><span class="sxs-lookup"><span data-stu-id="42137-157">Localization</span></span>

<span data-ttu-id="42137-p114">Некоторые элементы надстройки (например, имя, описание и загружаемый URL-адрес) необходимо локализовать для разных языковых стандартов. Эти элементы легко локализовать, указав значение по умолчанию и переопределения для языкового стандарта в дочернем элементе **Resources** элемента **VersionOverrides**. Ниже показано, как переопределить изображение, URL-адрес и строку.</span><span class="sxs-lookup"><span data-stu-id="42137-p114">Some aspects of the add-in need to be localized for different locales, such as the name, description and the URL that's loaded. These elements can easily be localized by specifying the default value and then locale overrides in the **Resources** element within the **VersionOverrides** element. The following shows how to override an image, a URL, and a string:</span></span>


```XML
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
      <!-- add information for other locales -->
    </bt:Image>
  </bt:Images>

  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
      <!-- add information for other locales -->
    </bt:Url>
  </bt:Urls>

  <bt:ShortStrings> 
    <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
      <bt:Override Locale="ar-sa" Value="<add localized value here>" />
      <!-- add information for other locales -->
    </bt:String>
  </bt:ShortStrings>
</Resources>
```

<span data-ttu-id="42137-161">Справочник по схеме содержит полные сведения о том, какие элементы можно локализовать.</span><span class="sxs-lookup"><span data-stu-id="42137-161">The schema reference contains full information on which elements can be localized.</span></span>

## <a name="hosts"></a><span data-ttu-id="42137-162">Hosts</span><span class="sxs-lookup"><span data-stu-id="42137-162">Hosts</span></span>

<span data-ttu-id="42137-163">Ниже показано, как указывается элемент **Hosts** для надстроек Outlook.</span><span class="sxs-lookup"><span data-stu-id="42137-163">Outlook add-ins specify the **Hosts** element like the following.</span></span>

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

<span data-ttu-id="42137-164">Он отличается от элемента **Hosts** в элементе **VersionOverrides**, который рассматривается в статье [Создание команд надстроек в манифесте для Excel, PowerPoint и Word](../develop/create-addin-commands.md).</span><span class="sxs-lookup"><span data-stu-id="42137-164">This is separate from the **Hosts** element inside the **VersionOverrides** element, which is discussed in [Create add-in commands in your manifest for Excel, PowerPoint, and Word](../develop/create-addin-commands.md).</span></span>

## <a name="requirements"></a><span data-ttu-id="42137-165">Требования</span><span class="sxs-lookup"><span data-stu-id="42137-165">Requirements</span></span>

<span data-ttu-id="42137-p115">Элемент **Requirements** указывает набор API-интерфейсов, доступный надстройке. Для надстройки Outlook требуются набор обязательных элементов Mailbox и версия 1.1 или выше. Последняя версия набора обязательных элементов указана в справочнике по API. Дополнительные сведения о наборах обязательных элементов см. в статье [API-интерфейсы Outlook](apis.md).</span><span class="sxs-lookup"><span data-stu-id="42137-p115">The **Requirements** element specifies the set of APIs available to the add-in. For an Outlook add-in, the requirement set must be Mailbox and a value of 1.1 or above. Please refer to the API reference for the latest requirement set version. Refer to the [Outlook add-in APIs](apis.md) for more information on requirement sets.</span></span>

<span data-ttu-id="42137-170">Элемент **Requirements** также может присутствовать в элементе **VersionOverrides**, позволяя надстройке указывать другие требования при загрузке в клиентах, поддерживающих **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="42137-170">The **Requirements** element can also appear in the **VersionOverrides** element, allowing the add-in to specify a different requirement when loaded in clients that support **VersionOverrides**.</span></span>

<span data-ttu-id="42137-171">В следующем примере используется атрибут **DefaultMinVersion** элемента **Sets**, чтобы запрашивался файл office.js версии 1.1 или выше, и атрибут **MinVersion** элемента **Set**, чтобы запрашивался набор требований Mailbox версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="42137-171">The following example uses the **DefaultMinVersion** attribute of the **Sets** element to require office.js version 1.1 or higher, and the **MinVersion** attribute of the **Set** element to require the Mailbox requirement set version 1.1.</span></span>

```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```

## <a name="form-settings"></a><span data-ttu-id="42137-172">Параметры формы</span><span class="sxs-lookup"><span data-stu-id="42137-172">Form settings</span></span>

<span data-ttu-id="42137-p116">Элемент **FormSettings** используется устаревшими клиентами Outlook, которые поддерживают только схему версии 1.1, а не **VersionOverrides**. С помощью этого элемента разработчики указывают, как надстройка будет отображаться в таких клиентах. Он состоит из двух частей: **ItemRead** и **ItemEdit**. **ItemRead** позволяет указать, как надстройка отображается при просмотре сообщений и встреч. **ItemEdit** описывает отображение надстройки при создании ответа, сообщения или встречи либо редактировании встречи организатором.</span><span class="sxs-lookup"><span data-stu-id="42137-p116">The **FormSettings** element is used by older Outlook clients, which only support schema 1.1 and not **VersionOverrides**. Using this element, developers define how the add-in will appear in such clients. There are two parts - **ItemRead** and **ItemEdit**. **ItemRead** is used to specify how the add-in appears when the user reads messages and appointments. **ItemEdit** describes how the add-in appears while the user is composing a reply, new message, new appointment or editing an appointment where they are the organizer.</span></span>

<span data-ttu-id="42137-p117">Эти параметры напрямую связаны с правилами активации в элементе **Rule**. Если надстройка указывает, что она должна отображаться на сообщении в форме создания, то должна быть указана форма **ItemEdit**.</span><span class="sxs-lookup"><span data-stu-id="42137-p117">These settings are directly related to the activation rules in the **Rule** element. For example, if an add-in specifies that it should appear on a message in compose mode, an **ItemEdit** form must be specified.</span></span>

<span data-ttu-id="42137-180">Дополнительные сведения см. в статье Schema reference for Office Add-ins manifests (v1.1).</span><span class="sxs-lookup"><span data-stu-id="42137-180">For more details, please refer to the [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>

## <a name="app-domains"></a><span data-ttu-id="42137-181">Домены приложений</span><span class="sxs-lookup"><span data-stu-id="42137-181">App domains</span></span>

<span data-ttu-id="42137-p118">Домен начальной страницы надстройки, заданной в элементе **SourceLocation**, является доменом по умолчанию для этой надстройки. Если элементы **AppDomains** и **AppDomain** не используются, а ваша надстройка попытается перейти к другому домену, в браузере откроется новое окно за пределами области надстройки. Чтобы надстройка могла переходить на другой домен в пределах области надстройки, добавьте элемент **AppDomains** и укажите каждый дополнительный домен в отдельном дочернем элементе **AppDomain** в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="42137-p118">The domain of the add-in start page that you specify in the **SourceLocation** element is the default domain for the add-in. Without using the **AppDomains** and **AppDomain** elements, if your add-in attempts to navigate to another domain, the browser will open a new window outside of the add-in pane. In order to allow the add-in to navigate to another domain within the add-in pane, add an **AppDomains** element and include each additional domain in its own **AppDomain** sub-element in the add-in manifest.</span></span>

<span data-ttu-id="42137-185">В следующем примере домен  `https://www.contoso2.com` указан как второй домен, к которому надстройка может переходить в рамках области надстройки:</span><span class="sxs-lookup"><span data-stu-id="42137-185">The following example specifies a domain  `https://www.contoso2.com` as a second domain that the add-in can navigate to within the add-in pane:</span></span>

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

<span data-ttu-id="42137-186">Домены надстроек также необходимы для обмена файлами cookie между всплывающим окном и надстройкой, запущенной в расширенном клиенте.</span><span class="sxs-lookup"><span data-stu-id="42137-186">App domains are also necessary to enable cookie sharing between the pop-out window and the add-in running in the rich client.</span></span>

<span data-ttu-id="42137-187">В следующей таблице описано поведение браузера при попытке перехода по URL-адресу за пределами стандартного домена надстройки.</span><span class="sxs-lookup"><span data-stu-id="42137-187">The following table describes browser behavior when your add-in attempts to navigate to a URL outside of the add-in's default domain.</span></span>

|<span data-ttu-id="42137-188">Клиент Outlook</span><span class="sxs-lookup"><span data-stu-id="42137-188">Outlook client</span></span>|<span data-ttu-id="42137-189">Домен определен</span><span class="sxs-lookup"><span data-stu-id="42137-189">Domain defined</span></span><br><span data-ttu-id="42137-190">в AppDomains?</span><span class="sxs-lookup"><span data-stu-id="42137-190">in AppDomains?</span></span>|<span data-ttu-id="42137-191">Поведение браузера</span><span class="sxs-lookup"><span data-stu-id="42137-191">Browser behavior</span></span>|
|---|---|---|
|<span data-ttu-id="42137-192">Все клиенты</span><span class="sxs-lookup"><span data-stu-id="42137-192">All clients</span></span>|<span data-ttu-id="42137-193">Да</span><span class="sxs-lookup"><span data-stu-id="42137-193">Yes</span></span>|<span data-ttu-id="42137-194">Ссылка откроется в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="42137-194">Link opens in add-in task pane.</span></span>|
|<span data-ttu-id="42137-195">Outlook 2016 для Windows (единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="42137-195">Outlook 2016 on Windows (one-time purchase)</span></span><br><span data-ttu-id="42137-196">Outlook 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="42137-196">Outlook 2013 on Windows</span></span>|<span data-ttu-id="42137-197">Нет</span><span class="sxs-lookup"><span data-stu-id="42137-197">No</span></span>|<span data-ttu-id="42137-198">Ссылка откроется в Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="42137-198">Link opens in Internet Explorer 11.</span></span>|
|<span data-ttu-id="42137-199">Другие клиенты</span><span class="sxs-lookup"><span data-stu-id="42137-199">Other clients</span></span>|<span data-ttu-id="42137-200">Нет</span><span class="sxs-lookup"><span data-stu-id="42137-200">No</span></span>|<span data-ttu-id="42137-201">Ссылка откроется в браузере пользователя, используемом по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="42137-201">Link opens in user's default browser.</span></span>|

<span data-ttu-id="42137-202">Дополнительные сведения см. в разделе [Укажите домены, которые необходимо открыть в окне надстройки](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).</span><span class="sxs-lookup"><span data-stu-id="42137-202">For more details, see the [Specify domains you want to open in the add-in window](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).</span></span>

## <a name="permissions"></a><span data-ttu-id="42137-203">Разрешения</span><span class="sxs-lookup"><span data-stu-id="42137-203">Permissions</span></span>

<span data-ttu-id="42137-p119">Элемент **Permissions** содержит необходимые надстройке разрешения. Как правило, следует указать минимальные необходимые разрешения, требуемые для надстройки, в зависимости от конкретных методов, которые вы собираетесь использовать. Например, для почтовой надстройки, которая активируется в форме создания и только считывает свойства элементов типа [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), но не записывает их и не вызывает метод [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) для получения доступа к любым операциям веб-служб Exchange, следует указать разрешение **ReadItem**. Дополнительные сведения о доступных разрешениях см. в статье [Указание разрешений для доступа надстройки Outlook к почтовому ящику пользователя](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="42137-p119">The **Permissions** element contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but does not write to item properties like [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and does not call [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to access any Exchange Web Services operations should specify **ReadItem** permission. For details on the available permissions, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

<span data-ttu-id="42137-208">**Четырехуровневая модель разрешений для почтовых надстроек**</span><span class="sxs-lookup"><span data-stu-id="42137-208">**Four-tier permissions model for mail add-ins**</span></span>

![Четырехуровневая модель разрешений для схемы почтовых приложений версии 1.1.](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a><span data-ttu-id="42137-210">Правила активации</span><span class="sxs-lookup"><span data-stu-id="42137-210">Activation rules</span></span>

<span data-ttu-id="42137-p120">Правила активации указываются в элементе **Rule**. Элемент **Rule** может отображаться как дочерний для элемента **OfficeApp** в манифестах версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="42137-p120">Activation rules are specified in the **Rule** element. The **Rule** element can appear as a child of the **OfficeApp** element in 1.1 manifests.</span></span>

<span data-ttu-id="42137-213">С помощью правил активации можно активировать надстройку при соблюдении одного или нескольких из представленных ниже условий в выбранном элементе.</span><span class="sxs-lookup"><span data-stu-id="42137-213">Activation rules can be used to activate an add-in based on one or more of the following conditions on the currently selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="42137-214">Правила активации применяются только к тем клиентам, которые не поддерживают элемент **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="42137-214">Activation rules only apply to clients that do not support the **VersionOverrides** element.</span></span>

- <span data-ttu-id="42137-215">Тип элемента и/или класс сообщения</span><span class="sxs-lookup"><span data-stu-id="42137-215">The item type and/or message class</span></span>

- <span data-ttu-id="42137-216">Наличие известной сущности определенного типа, например адреса или номера телефона</span><span class="sxs-lookup"><span data-stu-id="42137-216">The presence of a specific type of known entity, such as an address or phone number</span></span>

- <span data-ttu-id="42137-217">Совпадение с регулярным выражением в основном тексте, теме или электронном адресе отправителя</span><span class="sxs-lookup"><span data-stu-id="42137-217">A regular expression match in the body, subject, or sender email address</span></span>

- <span data-ttu-id="42137-218">Наличие вложения</span><span class="sxs-lookup"><span data-stu-id="42137-218">The presence of an attachment</span></span>

<span data-ttu-id="42137-219">Подробные сведения и примеры правил активации см. в статье [Правила активации для надстроек Outlook](activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="42137-219">For details and samples of activation rules, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>


## <a name="next-steps-add-in-commands"></a><span data-ttu-id="42137-220">Дальнейшие действия: команды надстроек</span><span class="sxs-lookup"><span data-stu-id="42137-220">Next steps: Add-in commands</span></span>

<span data-ttu-id="42137-p121">После определения основного манифеста определите команды для вашей надстройки. Команды надстроек представляют собой кнопки на ленте, с помощью которых пользователи могут легко и интуитивно активировать ваши надстройки. Дополнительные сведения см. в статье [Команды надстроек Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="42137-p121">After defining a basic manifest, define add-in commands for your add-in. Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="42137-224">Пример надстройки, в которой определены команды надстройки: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span><span class="sxs-lookup"><span data-stu-id="42137-224">For an example add-in that defines add-in commands, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span></span>

## <a name="next-steps-add-mobile-support"></a><span data-ttu-id="42137-225">Дальнейшие действия: Добавление поддержки мобильных устройств</span><span class="sxs-lookup"><span data-stu-id="42137-225">Next steps: Add mobile support</span></span>

<span data-ttu-id="42137-p122">При необходимости в надстройку можно добавить поддержку мобильной версии Outlook. Мобильная версия Outlook поддерживает команды надстроек примерно так же, как и Outlook для Windows и Mac. Дополнительные сведения см. в статье [Добавление поддержки команд надстроек для Outlook Mobile](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="42137-p122">Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="42137-229">См. также</span><span class="sxs-lookup"><span data-stu-id="42137-229">See also</span></span>

- [<span data-ttu-id="42137-230">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="42137-230">Localization for Office Add-ins</span></span>](../develop/localization.md)
- [<span data-ttu-id="42137-231">Конфиденциальность, разрешения и безопасность для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="42137-231">Privacy, permissions, and security for Outlook add-ins</span></span>](privacy-and-security.md)
- [<span data-ttu-id="42137-232">API надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="42137-232">Outlook add-in APIs</span></span>](apis.md)
- [<span data-ttu-id="42137-233">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="42137-233">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="42137-234">Справочник по схеме для манифестов надстроек Office (версия 1.1)</span><span class="sxs-lookup"><span data-stu-id="42137-234">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="42137-235">Оформление надстроек Office</span><span class="sxs-lookup"><span data-stu-id="42137-235">Design your Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="42137-236">Общие сведения о разрешениях для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="42137-236">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="42137-237">Использование правил активации на основе регулярных выражений для отображения надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="42137-237">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="42137-238">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="42137-238">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)