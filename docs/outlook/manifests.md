---
title: Манифесты надстройки Outlook
description: В манифесте описывается, как выполняется интеграция надстройки Outlook с клиентами Outlook, включая пример.
ms.date: 05/27/2020
localization_priority: Priority
ms.openlocfilehash: 8c5a31248f68e8f8b5b6ab4b2cf12c9bb969e062f0dccd68c8f5d7c3f5262452
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094196"
---
# <a name="outlook-add-in-manifests"></a>Манифесты надстройки Outlook

Надстройка Outlook состоит из двух компонентов: XML-манифеста надстройки и веб-страницы с поддержкой библиотеки JavaScript для надстроек Office (office.js). В манифесте описывается интеграция надстройки с клиентами Outlook. Ниже приведен пример.

 > [!NOTE]
 > Все значения URL-адресов в следующем примере начинаются со строки "https://appdemo.contoso.com". Это значение — заполнитель. В фактическом допустимом манифесте эти значения будут содержать действительные URL-адреса с префиксом HTTPS.

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

Не все клиенты Outlook поддерживают новейшие функции, а у некоторых пользователей Outlook установлена более ранняя версия Outlook. С помощью версий схемы разработчики могут создавать надстройки с обратной совместимостью, которые используют новые функции, если они доступны, но работают и в более ранних версиях.

Наглядный пример — элемент манифеста **VersionOverrides**. Все элементы, определенные в **VersionOverrides**, заменяют соответствующий элемент в другой части манифеста. Это означает, что по мере возможности Outlook будет использовать содержимое раздела **VersionOverrides** для установки параметров надстройки. Тем не менее, если версия Outlook не поддерживает определенную версию **VersionOverrides**, Outlook пропустит ее и будет использовать остальные сведения из манифеста. 

Этот подход означает, что разработчикам не требуется создавать несколько отдельных манифестов — все параметры можно задать в одном файле.

В настоящий момент доступны следующие версии схемы:


|Версия|Описание|
|:-----|:-----|
|v1.0|Поддерживает версию 1.0 API JavaScript для Office. Для надстроек Outlook поддерживается форма чтения. |
|1.1|Поддерживает версии 1.1 API JavaScript для Office и **VersionOverrides**. Для надстроек Outlook поддерживается форма создания.|
|**VersionOverrides** 1.0|Поддерживает более поздние версии API JavaScript для Office. Поддерживаются команды надстроек.|
|**VersionOverrides** 1.1|Поддерживает более поздние версии API JavaScript для Office. Поддерживает команды надстроек и добавляет поддержку новых функций, таких как [закрепляемые области задач](pinnable-taskpane.md) и мобильные надстройки.|

В этой статье рассматриваются требования для манифеста версии 1.1. Даже если в манифесте вашей надстройки используется элемент **VersionOverrides**, все равно важно включить элементы манифеста версии 1.1, чтобы надстройка работала со старыми клиентами, которые не поддерживают **VersionOverrides**.

> [!NOTE]
> Outlook использует схему для проверки манифестов. Поэтому элементы манифеста должны располагаться в определенном порядке. Если порядок не соблюден, при загрузке неопубликованной надстройки могут возникать ошибки. Вы можете скачать [определение схемы XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), которое позволит создать манифест с правильным расположением элементов.

## <a name="root-element"></a>Корневой элемент

Корневой элемент манифеста надстройки Outlook — **OfficeApp**. Этот элемент также объявляет пространство имен по умолчанию, версию схемы и тип надстройки. Поместите все остальные элементы манифеста между его открывающим и закрывающим тегами. Ниже приводится пример корневого элемента.


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

## <a name="version"></a>Version

Это версия конкретной надстройки. Когда разработчик обновляет какой-либо элемент манифеста, версию также необходимо увеличить. Таким образом, при установке нового манифеста заменяется имеющийся, а пользователю становятся доступны новые функции. Если эта надстройка была отправлена в магазин, манифест потребуется заново отправить и проверить. Спустя несколько часов (после утверждения обновленного манифеста) пользователи надстройки автоматически получат его.

При изменении разрешений, запрашиваемых надстройкой, пользователям предлагается выполнить обновление и повторно согласиться на предоставление надстройке разрешений. Если администратор установил эту надстройку для всей организации, он должен будет дать свое согласие первым. До этого пользователям будут доступны только старые функции.

## <a name="versionoverrides"></a>VersionOverrides

Элемент **VersionOverrides** — это расположение данных о [командах надстройки](add-in-commands-for-outlook.md).

Кроме того, в этом элементе определяется поддержка [мобильных надстроек](add-mobile-support.md).

Описание этого элемента см. в статье [Создание команд надстроек в манифесте для Excel, PowerPoint и Word](../develop/create-addin-commands.md).

## <a name="localization"></a>Локализация

Некоторые элементы надстройки (например, имя, описание и загружаемый URL-адрес) необходимо локализовать для разных языковых стандартов. Эти элементы легко локализовать, указав значение по умолчанию и переопределения для языкового стандарта в дочернем элементе **Resources** элемента **VersionOverrides**. Ниже показано, как переопределить изображение, URL-адрес и строку.


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

Ниже показано, как указывается элемент **Hosts** для надстроек Outlook.

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

Он отличается от элемента **Hosts** в элементе **VersionOverrides**, который рассматривается в статье [Создание команд надстроек в манифесте для Excel, PowerPoint и Word](../develop/create-addin-commands.md).

## <a name="requirements"></a>Требования

Элемент **Requirements** указывает набор API-интерфейсов, доступный надстройке. Для надстройки Outlook требуются набор обязательных элементов Mailbox и версия 1.1 или выше. Последняя версия набора обязательных элементов указана в справочнике по API. Дополнительные сведения о наборах обязательных элементов см. в статье [API-интерфейсы Outlook](apis.md).

Элемент **Requirements** также может присутствовать в элементе **VersionOverrides**, позволяя надстройке указывать другие требования при загрузке в клиентах, поддерживающих **VersionOverrides**.

В следующем примере используется атрибут **DefaultMinVersion** элемента **Sets**, чтобы запрашивался файл office.js версии 1.1 или выше, и атрибут **MinVersion** элемента **Set**, чтобы запрашивался набор требований Mailbox версии 1.1.

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

Элемент **FormSettings** используется устаревшими клиентами Outlook, которые поддерживают только схему версии 1.1, а не **VersionOverrides**. С помощью этого элемента разработчики указывают, как надстройка будет отображаться в таких клиентах. Он состоит из двух частей: **ItemRead** и **ItemEdit**. **ItemRead** позволяет указать, как надстройка отображается при просмотре сообщений и встреч. **ItemEdit** описывает отображение надстройки при создании ответа, сообщения или встречи либо редактировании встречи организатором.

Эти параметры напрямую связаны с правилами активации в элементе **Rule**. Если надстройка указывает, что она должна отображаться на сообщении в форме создания, то должна быть указана форма **ItemEdit**.

Дополнительные сведения см. в статье Schema reference for Office Add-ins manifests (v1.1).

## <a name="app-domains"></a>Домены приложений

Домен начальной страницы надстройки, заданной в элементе **SourceLocation**, является доменом по умолчанию для этой надстройки. Если элементы **AppDomains** и **AppDomain** не используются, а ваша надстройка попытается перейти к другому домену, в браузере откроется новое окно за пределами области надстройки. Чтобы надстройка могла переходить на другой домен в пределах области надстройки, добавьте элемент **AppDomains** и укажите каждый дополнительный домен в отдельном дочернем элементе **AppDomain** в манифесте надстройки.

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
|Outlook 2016 для Windows (единовременная покупка)<br>Outlook 2013 для Windows|Нет|Ссылка откроется в Internet Explorer 11.|
|Другие клиенты|Нет|Ссылка откроется в браузере пользователя, используемом по умолчанию.|

Дополнительные сведения см. в разделе [Укажите домены, которые необходимо открыть в окне надстройки](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).

## <a name="permissions"></a>Разрешения

Элемент **Permissions** содержит необходимые надстройке разрешения. Как правило, следует указать минимальные необходимые разрешения, требуемые для надстройки, в зависимости от конкретных методов, которые вы собираетесь использовать. Например, для почтовой надстройки, которая активируется в форме создания и только считывает свойства элементов типа [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), но не записывает их и не вызывает метод [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) для получения доступа к любым операциям веб-служб Exchange, следует указать разрешение **ReadItem**. Дополнительные сведения о доступных разрешениях см. в статье [Указание разрешений для доступа надстройки Outlook к почтовому ящику пользователя](understanding-outlook-add-in-permissions.md).

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

Правила активации указываются в элементе **Rule**. Элемент **Rule** может отображаться как дочерний для элемента **OfficeApp** в манифестах версии 1.1.

С помощью правил активации можно активировать надстройку при соблюдении одного или нескольких из представленных ниже условий в выбранном элементе.

> [!NOTE]
> Правила активации применяются только к тем клиентам, которые не поддерживают элемент **VersionOverrides**.

- Тип элемента и/или класс сообщения

- Наличие известной сущности определенного типа, например адреса или номера телефона

- Совпадение с регулярным выражением в основном тексте, теме или электронном адресе отправителя

- Наличие вложения

Подробные сведения и примеры правил активации см. в статье [Правила активации для надстроек Outlook](activation-rules.md).


## <a name="next-steps-add-in-commands"></a>Дальнейшие действия: команды надстроек

После определения основного манифеста определите команды для вашей надстройки. Команды надстроек представляют собой кнопки на ленте, с помощью которых пользователи могут легко и интуитивно активировать ваши надстройки. Дополнительные сведения см. в статье [Команды надстроек Outlook](add-in-commands-for-outlook.md).

Пример надстройки, в которой определены команды надстройки: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).

## <a name="next-steps-add-mobile-support"></a>Дальнейшие действия: Добавление поддержки мобильных устройств

При необходимости в надстройку можно добавить поддержку мобильной версии Outlook. Мобильная версия Outlook поддерживает команды надстроек примерно так же, как и Outlook для Windows и Mac. Дополнительные сведения см. в статье [Добавление поддержки команд надстроек для Outlook Mobile](add-mobile-support.md).

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