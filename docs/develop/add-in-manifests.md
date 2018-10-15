---
title: XML-манифест надстройки Office
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: 8d8363b80b948f30e13ccd8620178268e03f1d57
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505869"
---
# <a name="office-add-ins-xml-manifest"></a>XML-манифест надстройки Office

XML-файл манифеста надстройки Office описывает способ ее активации, когда пользователь устанавливает и использует эту надстройку для работы с документами и приложениями Office.

С помощью такого XML-файла манифеста надстройка Office может выполнять следующие действия:

* Предоставлять идентификатор, версию, описание, отображаемое имя и языковой стандарт по умолчанию.

* Указывать изображения, используемые для фирменной символики надстройки, и значки, используемые для [команд надстройки][] в ленте Office.

* Указывать, как надстройка интегрируется с Office, включая создаваемые ею элементы пользовательского интерфейса, например кнопки на ленте.

* Определять запрошенные размеры по умолчанию для надстроек содержимого, а также запрошенную высоту для надстроек Outlook.

* Объявлять разрешения, в которых нуждается Надстройка Office, например чтение или запись документа.

* В случае надстроек Outlook необходимо определить одно или несколько правил, указывающих контекст, в котором эти надстройки будут активироваться и взаимодействовать с сообщением, сведениями о встрече или приглашением на собрание.

> [!NOTE]
> Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).

## <a name="required-elements"></a>Обязательные элементы

В приведенной ниже таблице указаны обязательные элементы для трех типов надстроек Office.

### <a name="required-elements-by-office-add-in-type"></a>Обязательные элементы по типам надстроек Office

| Элемент                                                                                      | Содержимое | Область задач | Outlook |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| [OfficeApp][]                                                                                |    X    |     X     |    X    |
| [Id][]                                                                                       |    X    |     X     |    X    |
| [Версия][]                                                                                  |    X    |     X     |    X    |
| [ProviderName][]                                                                             |    X    |     X     |    X    |
| [DefaultLocale][]                                                                            |    X    |     X     |    X    |
| [DisplayName][]                                                                              |    X    |     X     |    X    |
| [Описание][]                                                                              |    X    |     X     |    X    |
| [IconUrl][]                                                                                  |    X    |     X     |    X    |
| [HighResolutionIconUrl][]                                                                    |    X    |     X     |    X    |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       |    X    |     X     |         |
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]                         |    X    |     X     |         |
| [DesktopSettings][]                                                                          |         |           |    X    |
| [SourceLocation (MailApp)][]                                                                 |         |           |    X    |
| [Разрешения (ContentApp)][]<br/>[Разрешения (TaskPaneApp)][]<br/>[Разрешения (MailApp)][] |    X    |     X     |    X    |
| [Правило (RuleCollection)][]<br/>[Правило (MailApp)][]                                             |         |           |    X    |
| [Требования (MailApp)*][]                                                                  |         |           |    X    |
| [Установка*][]<br/>[Установки (MailAppRequirements)*][]                                                 |         |           |    X    |
| [Форма*][]<br/>[FormSettings*][]                                                              |         |           |    X    |
| [Установка (требований)*][]                                                                     |    X    |     X     |         |
| [Хосты*][]                                                                                   |    X    |     X     |         |

_\*Элемент добавлен в схему манифеста для надстроек Office версии 1.1._

<!-- Links for above table -->

[officeapp]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/officeapp?view=office-js
[идентификатор]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id
[версия]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/version
[providername]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/providername
[defaultlocale]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultlocale
[displayname]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/displayname
[description]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/description
[iconurl]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/iconurl
[highresolutioniconurl]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/highresolutioniconurl
[defaultsettings (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultsettings
[sourcelocation (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: http://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[разрешения (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[разрешения (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[разрешения (mailapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[правило (rulecollection)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/rule
[правило  (mailapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/rule
[требования (mailapp)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements
[установка*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/set
[установки (mailapprequirements)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sets
[форма*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/form
[formsettings*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/formsettings
[Установка (требований)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sets
[хосты*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts

## <a name="hosting-requirements"></a>Требования к размещению

Все URI изображений, в частности используемых для [команд надстройки][], должны поддерживать кэширование. Сервер с изображением не должен возвращать заголовок `Cache-Control`, содержащий `no-cache`, `no-store` или подобные параметры в HTTP-отклике.

Все URL-адреса, например адреса исходных файлов, указанные в элементе [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation), должны быть **защищены с помощью SSL (HTTPS)**. [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a>Рекомендации по отправке решений в AppSource

Убедитесь, что идентификатор надстройки представляет собой допустимый и уникальный GUID. В Интернете доступно множество генераторов, с помощью которых можно создать уникальный GUID.

Надстройки, представленные в AppSource, также должны содержать элемент [SupportUrl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/supporturl). Для получения дополнительной информации см. [Политики проверки для приложений и надстроек, представленных в AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).

Чтобы указать домены, отличные от указанного в элементе [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) для сценариев проверки подлинности, используйте только элемент [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains).

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a>Указание доменов, которые необходимо открыть в окне надстройки

При работе в Office Online ваша панель задач может быть перенесена на любой URL-адрес. Однако на классических платформах, если надстройка пытается перейти на URL-адрес в домене, отличном от домена, где размещена начальная страница (как указано в элементе [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) файла манифеста), этот URL-адрес откроется в новом окне веб-обозревателя, а не в панели надстроек ведущего приложения Office.

Чтобы переопределить это поведение (Office настольного компьютера), укажите каждый домен, который вы хотите открыть, в окне надстройки в списке доменов, указанных в элементе [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) файла манифеста. Если надстройка пытается перейти к URL-адресу в домене, который находится в списке, он открывается в панели задач как в Office настольного ПК, так и в Office Online. Если он пытается перейти к URL-адресу, отсутствующему в списке, то в Office настольного ПК этот URL-адрес открывается в новом окне браузера (вне панели надстройки).

> [!NOTE]
> Это поведение относится только к корневой панели надстройки. Если на странице надстройки есть элемент iframe, он может быть перенаправлен на любой URL-адрес независимо от того, указан ли он в **AppDomains** – даже в классической версии Office.

В приведенном ниже примере XML-манифеста главная страница надстройки размещена в домене `https://www.contoso.com`, как указано в элементе **SourceLocation**. В нем также указан домен `https://www.northwindtraders.com` в элементе [AppDomain](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomain) из списка **AppDomains**. Страницы из домена www.northwindtraders.com будут открываться в области надстройки даже на ПК с Office.

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="manifest-v11-xml-file-examples-and-schemas"></a>XML-файлы манифеста версии 1.1: примеры и схемы
Ниже показаны примеры XML-файлов манифеста версии 1.1 для надстроек области задач, контентных надстроек и надстроек Outlook.

# <a name="task-panetabtabid-1"></a>[Область задач](#tab/tabid-1)

[Схема манифеста приложения панели задач](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

<!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

<!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

<!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
   <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
   <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
   <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currenly only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
            <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office Ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
                <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                 <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
                <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                     <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                     <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  - <Action xsi:type="ExecuteFunction">
                  <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                     <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
            <!-- Menu example -->
            <Control xsi:type="Menu" id="Contoso.Menu">
              <Label resid="Contoso.Dropdown.Label" />
              <Supertip>
                <Title resid="Contoso.Dropdown.Label" />
                <Description resid="Contoso.Dropdown.Tooltip" />
              </Supertip>
              <Icon>
                <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
              </Icon>
              <Items>
                <Item id="Contoso.Menu.Item1">
                  <Label resid="Contoso.Item1.Label"/>
                  <Supertip>
                    <Title resid="Contoso.Item1.Label" />
                    <Description resid="Contoso.Item1.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Item>

                <Item id="Contoso.Menu.Item2">
                  <Label resid="Contoso.Item2.Label"/>
                  <Supertip>
                    <Title resid="Contoso.Item2.Label" />
                    <Description resid="Contoso.Item2.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane2.Url" />
                  </Action>
                </Item>

              </Items>
            </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="contenttabtabid-2"></a>[Содержимое](#tab/tabid-2)

[Схема манифеста контентного приложения](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mailtabtabid-3"></a>[Почта](#tab/tabid-3)

[Схема манифеста почтового приложения](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">

  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Проверка манифеста и устранение связанных с ним неполадок

Сведения об устранении проблем, связанных с манифестом надстройки, см. в статье [Проверка манифеста и устранение связанных с ним неполадок](../testing/troubleshoot-manifest.md). Там указано, как проверить манифест согласно [XSD](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), а также как отладить манифест с помощью ведения журнала в среде выполнения.

## <a name="see-also"></a>См. также

* [Создание команд надстройки в манифесте][команды надстройки]
* [Указание ведущих приложений Office и требований API](specify-office-hosts-and-api-requirements.md)
* [Локализация надстроек Office](localization.md)
* [Справочная схема по манифестам надстроек для Office](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [Проверка манифеста и устранение связанных с ним неполадок](../testing/troubleshoot-manifest.md)

[команды надстройки]: create-addin-commands.md