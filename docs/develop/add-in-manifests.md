---
title: XML-манифест надстроек Office
description: Получите обзор манифеста надстройки Office и его использования.
ms.date: 09/28/2021
ms.localizationpriority: high
---

# <a name="office-add-ins-xml-manifest"></a>XML-манифест надстроек Office

XML-файл манифеста надстройки Office описывает способ ее активации, когда пользователь устанавливает и использует эту надстройку для работы с документами и приложениями Office.

С помощью XML-файла манифеста надстройка Office может выполнять следующие действия:

* предоставлять идентификатор, версию, описание, отображаемое имя и языковой стандарт по умолчанию.

* указывать изображения, используемые для фирменного оформления надстройки, и значки, используемые для [команд надстройки](create-addin-commands.md) на ленте приложения Office;

* указывать, как надстройка интегрируется с Office, включая создаваемые ею элементы пользовательского интерфейса, например кнопки на ленте;

* определять запрошенные размеры по умолчанию для контентных надстроек, а также запрошенную высоту для надстроек Outlook;

* объявлять разрешения, в которых нуждается надстройка Office, например чтение или запись документа;

* В случае надстроек Outlook необходимо определить одно или несколько правил, указывающих контекст, в котором эти надстройки будут активироваться и взаимодействовать с сообщением, сведениями о встрече или приглашением на собрание.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a>Обязательные элементы

В приведенной ниже таблице указаны обязательные элементы для трех типов надстроек Office.

> [!NOTE]
> Кроме того, есть обязательный порядок размещения элементов в родительском элементе. Дополнительные сведения см. в статье [Как определить правильный порядок элементов манифеста](manifest-element-ordering.md).

### <a name="required-elements-by-office-add-in-type"></a>Обязательные элементы по типам надстроек Office

| Элемент                                                                                      | Контентная | Для области задач | Outlook |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| [OfficeApp][]                                                                                |    X    |     X     |    X    |
| [Id][]                                                                                       |    X    |     X     |    X    |
| [Version][]                                                                                  |    X    |     X     |    X    |
| [ProviderName][]                                                                             |    X    |     X     |    X    |
| [DefaultLocale][]                                                                            |    X    |     X     |    X    |
| [DisplayName][]                                                                              |    X    |     X     |    X    |
| [Description][]                                                                              |    X    |     X     |    X    |
| [IconUrl][]                                                                                  |    X    |     X     |    X    |
| [SupportUrl][]\*\*                                                                           |    X    |     X     |    X    |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       |    X    |     X     |         |
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]                         |    X    |     X     |         |
| [DesktopSettings][]                                                                          |         |           |    X    |
| [SourceLocation (MailApp)][]                                                                 |         |           |    X    |
| [Permissions (ContentApp)][]<br/>[Permissions (TaskPaneApp)][]<br/>[Permissions (MailApp)][] |    X    |     X     |    X    |
| [Rule (RuleCollection)][]<br/>[Rule (MailApp)][]                                             |         |           |    X    |
| [Requirements (MailApp)*][]                                                                  |         |           |    X    |
| [Set*][]<br/>[Sets (MailAppRequirements)*][]                                                 |         |           |    X    |
| [Form*][]<br/>[FormSettings*][]                                                              |         |           |    X    |
| [Sets (Requirements)*][]                                                                     |    X    |     X     |         |
| [Hosts*][]                                                                                   |    X    |     X     |         |

_\*Элемент добавлен в схеме манифеста для надстроек Office версии 1.1._

_\*\* SupportUrl требуется только для надстроек распространяемых с помощью AppSource._

<!-- Links for above table -->

[officeapp]: ../reference/manifest/officeapp.md
[id]: ../reference/manifest/id.md
[version]: ../reference/manifest/version.md
[providername]: ../reference/manifest/providername.md
[defaultlocale]: ../reference/manifest/defaultlocale.md
[displayname]: ../reference/manifest/displayname.md
[description]: ../reference/manifest/description.md
[iconurl]: ../reference/manifest/iconurl.md
[supporturl]: ../reference/manifest/supporturl.md
[defaultsettings (contentapp)]: ../reference/manifest/defaultsettings.md
[defaultsettings (taskpaneapp)]: ../reference/manifest/defaultsettings.md
[sourcelocation (contentapp)]: ../reference/manifest/sourcelocation.md
[sourcelocation (taskpaneapp)]: ../reference/manifest/sourcelocation.md
[desktopsettings]: /previous-versions/office/fp179684%28v=office.15%29
[sourcelocation (mailapp)]: /previous-versions/office/fp123668%28v=office.15%29
[permissions (contentapp)]: ../reference/manifest/permissions.md
[permissions (taskpaneapp)]: ../reference/manifest/permissions.md
[permissions (mailapp)]: ../reference/manifest/permissions.md
[rule (rulecollection)]: ../reference/manifest/rule.md
[rule (mailapp)]: ../reference/manifest/rule.md
[Requirements (MailApp)*]: ../reference/manifest/requirements.md
[Set*]: ../reference/manifest/set.md
[Sets (MailAppRequirements)*]: ../reference/manifest/sets.md
[Form*]: ../reference/manifest/form.md
[FormSettings*]: ../reference/manifest/formsettings.md
[Sets (Requirements)*]: ../reference/manifest/sets.md
[Hosts*]: ../reference/manifest/hosts.md

## <a name="hosting-requirements"></a>Требования к размещению

Все URI изображений, в частности используемые для [команд надстройки](create-addin-commands.md), должны поддерживать кэширование. Сервер с изображением не должен возвращать заголовок `Cache-Control`, содержащий `no-cache`, `no-store` или подобные параметры в ответе HTTP.

Все URL-адреса, например адреса исходных файлов, указанные в элементе [SourceLocation](../reference/manifest/sourcelocation.md), должны быть **защищены с помощью SSL (HTTPS)**. [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a>Рекомендации по отправке решений в AppSource

Убедитесь, что идентификатор надстройки представляет собой допустимый и уникальный GUID. В Интернете доступно множество генераторов, с помощью которых можно создать уникальный GUID.

Надстройки, отправляемые в AppSource, также должны включать элемент [SupportUrl](../reference/manifest/supporturl.md). Дополнительные сведения см. в статье [Политики проверки для приложений и надстроек, отправляемых в AppSource](/legal/marketplace/certification-policies).

Чтобы указать домены, отличные от указанного в элементе [SourceLocation](../reference/manifest/sourcelocation.md) для сценариев проверки подлинности, используйте только элемент [AppDomains](../reference/manifest/appdomains.md).

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a>Укажите домены, которые необходимо открыть в окне надстройки

В Office в Интернете область задач может открывать любой URL-адрес. Если на платформах для настольных компьютеров надстройка пытается перейти на URL-адрес в домене, отличном от домена, где размещена начальная страница (указан в элементе [SourceLocation](../reference/manifest/sourcelocation.md) файла манифеста), этот URL-адрес откроется в новом окне браузера, а не в области надстроек приложения Office.

Чтобы переопределить это поведение, укажите все домены, которые должны открываться в окне надстройки, в списке доменов в элементе [AppDomains](../reference/manifest/appdomains.md) файла манифеста. URL-адреса в доменах из списка будут открываться в области задач как в классическом Office, так и в Office в Интернете. URL-адреса в доменах не из списка будут открываться в новом окне браузера (не в области надстроек) в классическом Office.

> [!NOTE]
> Из этого правила есть два исключения.
>
> - Это относится только к корневой области надстройки. Если в страницу надстройки внедрен iframe, его можно перенаправить на любой URL-адрес, независимо от того, указан ли он в элементе **AppDomains**, даже в классической версии Office.
> - Если диалоговое окно открыто с помощью API [displayDialogAsync](/javascript/api/office/office.ui?view=common-js&preserve-view=true#office-office-ui-displaydialogasync-member(1)), URL-адрес, передаваемый методу, должен находиться в том же домене, что и надстройка. Затем диалоговое окно можно перенаправить на любой URL-адрес, независимо от того, указан ли он в элементе **AppDomains**, даже в классической версии Office.

В приведенном ниже примере XML-манифеста главная страница надстройки размещена в домене `https://www.contoso.com`, как указано в элементе **SourceLocation**. Кроме того, указан домен `https://www.northwindtraders.com` с помощью элемента [AppDomain](../reference/manifest/appdomain.md) из списка **AppDomains**. Если надстройка выполняет переход на страницу в домене `www.northwindtraders.com`, эта страница открывается в области надстроек (даже в классической версии Office).

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="version-overrides-in-the-manifest"></a>Переопределение версии в манифесте

Необязательный элемент [VersionOverrides](../reference/manifest/versionoverrides.md) заслуживает особого упоминания. Он содержит вложенную дочернюю разметку, которая активирует дополнительные функции надстройки. Вот некоторые из этих функций:

 - Настройка ленты и меню Office.
 - Настройка работы Office со встроенной средой времени выполнения браузера, в которой запускаются надстройки.
 - Настройка взаимодействия надстройки с Azure Active Directory и Microsoft Graph для единого входа в систему.

Некоторые дочерние элементы `VersionOverrides` имеют значения, переопределяющие значения родительского `OfficeApp` элемента. Например, элемент `Hosts` в `VersionOverrides` переопределяет элемент `Hosts` в `OfficeApp`.

Элемент `VersionOverrides` имеет собственную схему, точнее — четыре схемы, в зависимости от типа надстройки и используемых функций. Вот эти схемы:

- [Область задач 1.0](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)
- [Контент 1.0](/openspecs/office_file_formats/ms-owemxml/c9cb8dca-e9e7-45a7-86b7-f1f0833ce2c7)
- [Почта 1.0](/openspecs/office_file_formats/ms-owemxml/578d8214-2657-4e6a-8485-25899e772fac)
- [Почта 1.1](/openspecs/office_file_formats/ms-owemxml/8e722c85-eb78-438c-94a4-edac7e9c533a)

Когда используется элемент `VersionOverrides`, элемент `OfficeApp` должен иметь атрибут `xmlns`, который определяет соответствующую схему. Возможные значения атрибута:

- `http://schemas.microsoft.com/office/taskpaneappversionoverrides`
- `http://schemas.microsoft.com/office/contentappversionoverrides`
- `http://schemas.microsoft.com/office/mailappversionoverrides`

Сам элемент `VersionOverrides` также должен иметь атрибут `xmlns`, определяющий схему. Возможные значения: три, приведенные выше, а также следующие:

- `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`

Элемент `VersionOverrides` также должен иметь атрибут `xsi:type`, который определяет версию схемы. Возможные значения:

- `VersionOverridesV1_0`
- `VersionOverridesV1_1`

Ниже приводятся примеры использования `VersionOverrides` в надстройке области задач и в надстройке почты соответственно. Обратите внимание, что когда используется почта `VersionOverrides` с версией 1.1, она должна быть последним дочерним элементом родительского элемента `VersionOverrides` типа 1.0. Значения дочерних элементов во внутреннем `VersionOverrides` переопределяют значения одноименных элементов в родительском вложении `VersionOverrides` и прародительском `OfficeApp` элементе.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <!-- child elements omitted -->
</VersionOverrides>
```

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <!-- other child elements omitted -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <!-- child elements omitted -->
  </VersionOverrides>
</VersionOverrides>
```

Пример манифеста, включающего элемент `VersionOverrides`, см. в примерах[ и схемах](#manifest-v11-xml-file-examples-and-schemas) XML-файлов Манифеста версии 1.1.

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a>Указание доменов, из которых выполняются вызовы API Office.js

Ваша надстройка может выполнять вызовы API Office.js из домена, указанного в элементе [SourceLocation](../reference/manifest/sourcelocation.md) файла манифеста. Если в вашей надстройке есть другие блоки IFrame, которым требуется доступ к API Office.js, добавьте домен этого исходного URL-адреса в список, указанный в элементе [AppDomains](../reference/manifest/appdomains.md) файла манифеста. Если блок IFrame с источником, не содержащимся в списке `AppDomains`, попытается выполнить вызов API Office.js, надстройка получит [ошибку об отказе в разрешении](../reference/javascript-api-for-office-error-codes.md).

## <a name="manifest-v11-xml-file-examples-and-schemas"></a>XML-файлы манифеста версии 1.1: примеры и схемы

Ниже показаны примеры XML-файлов манифеста версии 1.1 для надстроек области задач, контентных надстроек и надстроек Outlook.

# <a name="task-pane"></a>[Области задач](#tab/tabid-1)

[Схемы манифестов надстроек](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office app ribbon-->
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
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://myCDN/Images/ButtonFunction.png" />
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

# <a name="content"></a>[Контент](#tab/tabid-2)

[Схемы манифестов надстроек](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

# <a name="mail"></a>[почта](#tab/tabid-3);

[Схемы манифестов надстроек](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
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
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

## <a name="validate-an-office-add-ins-manifest"></a>Проверка манифеста надстройки Office

Сведения о проверке манифеста с помощью [определения схемы XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) см. в статье [Проверка манифеста надстройки Office](../testing/troubleshoot-manifest.md).

## <a name="see-also"></a>См. также

* [Определение правильного порядка элементов манифеста](manifest-element-ordering.md)
* [Создание команд надстройки в манифесте](create-addin-commands.md)
* [Указание приложений Office и обязательных элементов API](specify-office-hosts-and-api-requirements.md)
* [Локализация надстроек для Office](localization.md)
* [Справочная схема по манифестам надстроек для Office](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
* [Обновление API и версии манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md)
* [Определение аналогичной надстройки COM](make-office-add-in-compatible-with-existing-com-add-in.md)
* [Запрос разрешений на использование API в надстройках](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
* [Проверка манифеста надстройки Office](../testing/troubleshoot-manifest.md)
