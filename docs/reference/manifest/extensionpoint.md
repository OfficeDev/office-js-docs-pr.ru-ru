---
title: Элемент ExtensionPoint в файле манифеста
description: Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 44824e0c74b35105833f1f05cdda87bc873a4427
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094458"
---
# <a name="extensionpoint-element"></a>Элемент ExtensionPoint

 Определяет, где доступны функции надстройки в пользовательском интерфейсе Office. Элемент **ExtensionPoint** является дочерним для элемента [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Да  | Тип определяемой точки расширения.|

## <a name="extension-points-for-excel-only"></a>Точки расширения только для Excel

- **CustomFunctions** — пользовательская функция, написанная на JavaScript для Excel.

[В этом примере кода XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) показано, как использовать элемент **ExtensionPoint** со значением атрибута **CustomFunctions** и какие дочерние элементы следует использовать.

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Точки расширения для команд надстроек Word, Excel, PowerPoint и OneNote

- **PrimaryCommandSurface** — лента в Office.
- **ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в интерфейсе Office.

В приведенных ниже примерах показано, как применять элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.

> [!IMPORTANT]
> For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname">

```XML
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

#### <a name="child-elements"></a>Дочерние элементы
 
|**Элемент**|**Описание**|
|:-----|:-----|
|**CustomTab**|Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.|
|**OfficeTab**|Является обязательным, если вы хотите расширить вкладку ленты приложения Office по умолчанию (с помощью **PrimaryCommandSurface**). Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно. Дополнительные сведения см. в разделе [OfficeTab](officetab.md).|
|**OfficeMenu**|Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: <br/> - **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> - **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.|
|**Group**|A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.|
|**Label**|Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.|
|**Icon**|Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.|
|**Tooltip**|Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.|
|**Control**|В каждой группе должен быть по крайней мере один элемент управления. Элемент **управления** может быть либо **кнопкой** , либо **меню**. Используйте **меню** , чтобы указать раскрывающийся список элементов управления "Кнопка". В настоящее время поддерживаются только кнопки и меню. Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](control.md#button-control) и [Элементы управления меню](control.md#menu-dropdown-button-controls).<br/>**Примечание:**  Чтобы упростить устранение неполадок, рекомендуется добавлять элемент **Control** и соответствующие дочерние элементы **Resources** по одному.|
|**Script**|Ссылка на файл JavaScript с пользовательским определением функции и кодом регистрации. Этот элемент не используется в предварительной версии для разработчиков. Загрузку всех файлов JavaScript выполняет страница HTML.|
|**Page**|Ссылка на HTML-страницу для пользовательских функций.|

## <a name="extension-points-for-outlook"></a>Точки расширения для Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (можно использовать только в [DesktopFormFactor](desktopformfactor.md))
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [мобилеонлинемитингкоммандсурфаце](#mobileonlinemeetingcommandsurface-preview)
- [LaunchEvent](#launchevent-preview)
- [Events](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface

This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

#### <a name="officetab-example"></a>Пример элемента OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface

Эта точка расширения добавляет кнопки на ленту для надстроек, использующих форму создания сообщения. 

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

#### <a name="officetab-example"></a>Пример элемента OfficeTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для организатора собрания. 

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

#### <a name="officetab-example"></a>Пример элемента OfficeTab

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для участника собрания. 

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

#### <a name="officetab-example"></a>Пример элемента OfficeTab

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

Эта точка расширения добавляет кнопки на ленту для расширения модуля.

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface

Эта точка расширения помещает кнопки на панель команд для чтения почты в форм-факторе мобильного устройства.

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [Group](group.md) |  Добавляет группу кнопок на панель команд.  |

У элементов **ExtensionPoint** этого типа может быть только один дочерний элемент **Group**.

Для атрибута **xsi:type** элементов **Control**, содержащихся в этой точке расширения, должно быть назначено значение `MobileButton`.

#### <a name="example"></a>Пример

```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface-preview"></a>Мобилеонлинемитингкоммандсурфаце (Предварительная версия)

> [!NOTE]
> Эта точка расширения поддерживается только в [предварительной версии](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Android с подпиской на Microsoft 365.

Эта точка расширения помещает переключатель, подходящий для режима, на поверхности команды для встречи в мобильном конструктивном параметре. Организатор собрания может создать собрание по сети. Затем участник может присоединиться к собранию по сети. Чтобы узнать больше об этом сценарии, ознакомьтесь со статьей [Создание надстройки Outlook Mobile для веб-службы "поставщик собраний](../../outlook/online-meeting.md) ".

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [Control](control.md) |  Добавляет кнопку на поверхность команды.  |

`ExtensionPoint`у элементов этого типа может быть только один дочерний элемент: `Control` element.

`Control`Атрибуту элемента, содержащегося в этой точке расширения, должен быть `xsi:type` присвое значение `MobileButton` .

`Icon`Изображения должны быть в градациях серого с использованием шестнадцатеричного кода `#919191` или его эквивалента в [других цветовых форматах](https://convertingcolors.com/hex-color-919191.html).

#### <a name="example"></a>Пример

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
    <Label resid="residUILessButton0Name" />
    <Icon>
      <bt:Image resid="UiLessIcon" size="25" scale="1" />
      <bt:Image resid="UiLessIcon" size="25" scale="2" />
      <bt:Image resid="UiLessIcon" size="25" scale="3" />
      <bt:Image resid="UiLessIcon" size="32" scale="1" />
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="32" scale="3" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>insertContosoMeeting</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### <a name="launchevent-preview"></a>Лаунчевент (Предварительная версия)

> [!NOTE]
> Эта точка расширения поддерживается только в [предварительном просмотре](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) в Outlook в Интернете с подпиской на Microsoft 365.

Эта точка расширения позволяет активировать надстройку на основе поддерживаемых событий на настольных формах. В настоящее время единственными поддерживаемыми событиями являются `OnNewMessageCompose` и `OnNewAppointmentOrganizer` . Чтобы узнать больше об этом сценарии, ознакомьтесь со статьей [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md) .

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  Список [лаунчевент](launchevent.md) для активации на основе событий.  |
| [SourceLocation](sourcelocation.md) |  Расположение исходного файла JavaScript.  |

#### <a name="example"></a>Пример

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

### <a name="events"></a>События

Эта точка расширения добавляет обработчик для указанного события. Для получения дополнительных сведений об использовании этой точки расширения, ознакомьтесь со статьей [функция On Send для надстроек Outlook](../../outlook/outlook-on-send-addins.md).

| Элемент | Описание  |
|:-----|:-----|
|  [Event](event.md) |  Задает событие и функцию его обработчика.  |

#### <a name="itemsend-event-example"></a>Пример события ItemSend

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a>DetectedEntity

Эта точка расширения добавляет активацию контекстной надстройки для указанного типа сущности.

В соответствующем элементе [VersionOverrides](versionoverrides.md) для атрибута `xsi:type` должно быть задано значение `VersionOverridesV1_1`.

> [!NOTE]
> Этот тип элемента доступен в [клиентах Outlook, поддерживающих наборы обязательных требований 1.6 и более поздних версий.](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)

|  Элемент |  Описание  |
|:-----|:-----|
|  [Label](#label) |  Задает метку для надстройки в контекстном окне.  |
|  [SourceLocation](sourcelocation.md) |  Задает URL-адрес контекстного окна.  |
|  [Rule](rule.md) |  Задает одно или несколько правил, определяющих, когда активируется надстройка.  |

#### <a name="label"></a>Label

Обязательный элемент. Метка группы. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .

#### <a name="highlight-requirements"></a>Требования к выделению

The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.

However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.

- Сущности `EmailAddress` и `Url` не поддерживают выделение, поэтому их нельзя использовать для активации надстройки.
- Если используется одно правило, то для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.
- Если используется правило `RuleCollection`, совмещенное с другими правилами с помощью оператора `Mode="AND"`, то как минимум в одном из правил для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.
- Если используется правило `RuleCollection`, в котором правила совмещаются с помощью оператора `Mode="OR"`, то в каждом из них для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.

#### <a name="detectedentity-event-example"></a>Пример события DetectedEntity

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint>
```
