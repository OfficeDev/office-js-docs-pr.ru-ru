---
title: Манифесты надстройки Outlook
description: Общие сведения о двух типах манифестов, доступных для надстроек Outlook.
ms.date: 10/18/2022
ms.localizationpriority: high
ms.openlocfilehash: a22b5180fee6b4f9f0663eff54b57510016202a2
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607557"
---
# <a name="outlook-add-in-manifests"></a>Манифесты надстройки Outlook

Надстройка Outlook состоит из двух компонентов: манифеста надстройки и веб-приложения, поддерживаемого библиотекой JavaScript для надстроек Office (office.js). В манифесте описывается, как надстройка интегрируется с клиентами Outlook.

Существует два возможных формата манифеста: XML и JSON. Вы можете узнать о манифесте JSON в [манифесте Teams для надстроек Office (предварительная версия)](../develop/json-manifest-overview.md). Эта статья посвящена XML-манифесту.

Ниже приведен пример XML-манифеста.

 > [!NOTE]
 > All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.

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

## <a name="schema-versions"></a>Версии схемы

Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.

Примером этого является элемент **\<VersionOverrides\>** в манифесте. Все элементы, определенные в **\<VersionOverrides\>**, заменяют соответствующий элемент в другой части манифеста. Это означает, что по мере возможности Outlook будет использовать содержимое раздела **\<VersionOverrides\>** для установки параметров надстройки. Тем не менее, если версия Outlook не поддерживает определенную версию **\<VersionOverrides\>**, Outlook пропустит ее и будет использовать остальные сведения из манифеста. 

Этот подход означает, что разработчикам не требуется создавать несколько отдельных манифестов — все параметры можно задать в одном файле.

В настоящий момент доступны следующие версии схемы:


|Версия|Описание|
|:-----|:-----|
|v1.0|Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form. |
|1.1|Поддерживает версию 1.1 API JavaScript для Office и **\<VersionOverrides\>**. Для надстроек Outlook поддерживается форма создания.|
|**\<VersionOverrides\>** 1.0|Поддерживает более поздние версии API JavaScript для Office. Поддерживаются команды надстроек.|
|**\<VersionOverrides\>** 1.1|Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.|

В этой статье рассматриваются требования для манифеста версии 1.1. Даже если в манифесте вашей надстройки используется элемент **\<VersionOverrides\>**, все равно важно включить элементы манифеста версии 1.1, чтобы надстройка работала со старыми клиентами, которые не поддерживают **\<VersionOverrides\>**.

> [!NOTE]
> Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.

## <a name="root-element"></a>Корневой элемент

Корневой элемент манифеста надстройки Outlook — **\<OfficeApp\>**. Этот элемент также объявляет пространство имен по умолчанию, версию схемы и тип надстройки. Поместите все остальные элементы манифеста между его открывающим и закрывающим тегами. Ниже приводится пример корневого элемента.


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

## <a name="version"></a>Версия

This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.

If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.

## <a name="versionoverrides"></a>VersionOverrides

Элемент **\<VersionOverrides\>** — это расположение данных о [командах надстройки](add-in-commands-for-outlook.md).

Кроме того, в этом элементе определяется поддержка [мобильных надстроек](add-mobile-support.md).

Описание этого элемента см. в статье [Создание команд надстроек в манифесте для Excel, PowerPoint и Word](../develop/create-addin-commands.md).

## <a name="localization"></a>Локализация

Некоторые элементы надстройки (например, имя, описание и загружаемый URL-адрес) необходимо локализовать для разных языковых стандартов. Эти элементы легко локализовать, указав значение по умолчанию и переопределения для языкового стандарта в элементе **\<Resources\>** элемента **\<VersionOverrides\>**. Ниже показано, как переопределить изображение, URL-адрес и строку.


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

Справочник по схеме содержит полные сведения о том, какие элементы можно локализовать.

## <a name="hosts"></a>Hosts

Ниже показано, как указывается элемент **\<Hosts\>** для надстроек Outlook.

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

Он отличается от элемента **\<Hosts\>** в элементе **\<VersionOverrides\>**, который рассматривается в статье [Создание команд надстроек в манифесте для Excel, PowerPoint и Word](../develop/create-addin-commands.md).

## <a name="requirements"></a>Требования

Элемент **\<Requirements\>** указывает набор API-интерфейсов, доступный надстройке. Для надстройки Outlook требуются набор требований Mailbox и версия 1.1 или выше. Последняя версия набора требований указана в справочнике по API. Дополнительные сведения о наборах обязательных элементов см. в статье [API надстроек Outlook](apis.md).

Элемент **\<Requirements\>** также может присутствовать в элементе **\<VersionOverrides\>**, позволяя надстройке указывать другие требования при загрузке в клиентах, поддерживающих **\<VersionOverrides\>**.

В следующем примере используется атрибут **DefaultMinVersion** элемента **\<Sets\>**, чтобы запрашивался файл office.js версии 1.1 или выше, и атрибут **MinVersion** элемента **\<Set\>**, чтобы запрашивался набор обязательных элементов Mailbox версии 1.1.

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

## <a name="form-settings"></a>Параметры формы

Элемент **\<FormSettings\>** используется более ранними клиентами Outlook, которые поддерживают только схему версии 1.1, а не **\<VersionOverrides\>**. С помощью этого элемента разработчики указывают, как надстройка будет отображаться в таких клиентах. Он состоит из двух частей: **ItemRead** и **ItemEdit**. **ItemRead** позволяет указать, как надстройка отображается при просмотре сообщений и встреч. **ItemEdit** описывает отображение надстройки при создании ответа, сообщения или встречи либо редактировании встречи организатором.

Эти параметры напрямую связаны с правилами активации в элементе **\<Rule\>**. Если надстройка указывает, что она должна отображаться на сообщении в форме создания, то должна быть указана форма  **ItemEdit**.

Дополнительные сведения см. в статье Schema reference for Office Add-ins manifests (v1.1).

## <a name="app-domains"></a>Домены приложений

Домен начальной страницы надстройки, заданной в элементе **\<SourceLocation\>**, является доменом по умолчанию для этой надстройки. Если элементы **\<AppDomains\>** и **\<AppDomain\>** не используются, а ваша надстройка попытается перейти к другому домену, в браузере откроется новое окно за пределами области надстройки. Чтобы надстройка могла переходить на другой домен в пределах области надстройки, добавьте элемент **\<AppDomains\>** и укажите каждый дополнительный домен в отдельном дочернем элементе **\<AppDomain\>** в манифесте надстройки.

В следующем примере домен  `https://www.contoso2.com` указан как второй домен, к которому надстройка может переходить в рамках области надстройки:

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

Домены надстроек также необходимы для обмена файлами cookie между всплывающим окном и надстройкой, запущенной в расширенном клиенте.

В следующей таблице описано поведение браузера при попытке перехода по URL-адресу за пределами стандартного домена надстройки.

|Клиент Outlook|Домен определен<br>в AppDomains?|Поведение браузера|
|---|---|---|
|Все клиенты|Да|Ссылка откроется в области задач надстройки.|
|— Outlook 2016 Windows (бессрочное использование с корпоративным лицензированием)<br>— Outlook 2013 в Windows (бессрочно)|Нет|Ссылка откроется в Internet Explorer 11.|
|Другие клиенты|Нет|Ссылка откроется в браузере пользователя, используемом по умолчанию.|

Дополнительные сведения см. в разделе [Укажите домены, которые необходимо открыть в окне надстройки](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).

## <a name="permissions"></a>Разрешения

Элемент **\<Permissions\>** содержит необходимые надстройке разрешения. В целом вам следует указать минимальные необходимые разрешения, требуемые для вашей надстройки, в зависимости от конкретных методов, которые вы собираетесь использовать. Например, для почтовой надстройки, которая активируется в форме создания и только считывает свойства элементов типа [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), но не записывает их и не вызывает метод [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) для получения доступа к любым операциям веб-служб Exchange, следует указать разрешение **ReadItem**. Дополнительные сведения о доступных разрешениях см. в разделе [Общие сведения о разрешениях для надстроек Outlook](understanding-outlook-add-in-permissions.md).

**Четырехуровневая модель разрешений для почтовых надстроек**

![Четырехуровневая модель разрешений для схемы почтовых приложений версии 1.1.](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a>Правила активации

Правила активации указываются в элементе **\<Rule\>**. Элемент **\<Rule\>** может отображаться как дочерний для элемента **\<OfficeApp\>** в манифестах версии 1.1.

С помощью правил активации можно активировать надстройку при соблюдении одного или нескольких из представленных ниже условий в выбранном элементе.

> [!NOTE]
> Правила активации применяются только к тем клиентам, которые не поддерживают элемент **\<VersionOverrides\>**.

- Тип элемента и/или класс сообщения

- Наличие известной сущности определенного типа, например адреса или номера телефона

- Совпадение с регулярным выражением в основном тексте, теме или электронном адресе отправителя

- Наличие вложения

Подробные сведения и примеры правил активации см. в статье [Правила активации для надстроек Outlook](activation-rules.md).

## <a name="next-steps-add-in-commands"></a>Дальнейшие действия: команды надстроек

After defining a basic manifest, define add-in commands for your add-in. Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

Пример надстройки, в которой определены команды надстройки: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).

## <a name="next-steps-add-mobile-support"></a>Дальнейшие действия: Добавление поддержки мобильных устройств

Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).

## <a name="see-also"></a>См. также

- [Локализация надстроек для Office](../develop/localization.md)
- [Конфиденциальность, разрешения и безопасность для надстроек Outlook](privacy-and-security.md)
- [API надстроек Outlook](apis.md)
- [XML-манифест надстройки Office](../develop/add-in-manifests.md)
- [Справочник по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md)
- [Оформление надстроек Office](../design/add-in-design.md)
- [Общие сведения о разрешениях для надстроек Outlook](understanding-outlook-add-in-permissions.md)
- [Использование правил активации на основе регулярных выражений для отображения надстройки Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md)