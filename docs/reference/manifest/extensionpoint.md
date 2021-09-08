---
title: Элемент ExtensionPoint в файле манифеста
description: Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 8f84be1f2dcc43d795026fcd28dc3860c5e07a1e
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937542"
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
> Для элементов, которые содержат атрибут ID, обязательно предоставляйте уникальный идентификатор. Мы рекомендуем использовать название вашей компании и личный идентификатор. Пример формата приведен ниже. <CustomTab id="mycompanyname.mygroupname">

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
 
|Элемент|Описание|
|:-----|:-----|
|**CustomTab**|Обязательный, если требуется добавить пользовательскую вкладку в ленту (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **CustomTab** и **OfficeTab** одновременно. Атрибут **id** является обязательным. |
|**OfficeTab**|Требуется, если вы хотите расширить вкладку ленты Приложение Office по умолчанию (с помощью **PrimaryCommandSurface).** Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно. Подробные сведения см. [в материале OfficeTab](officetab.md).|
|**OfficeMenu**|Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: <br/> - **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши. <br/> - **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь нажимает ячейку электронной таблицы правой кнопкой мыши.|
|**Group**|Группа точек расширения интерфейса пользователя на вкладке. В группе может быть до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов. |
|**Label**|Обязательный. Метка группы. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String.** **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**.|
|**Icon**|Обязательный. Определяет значок группы для использования на устройствах с малым форм-фактором или в случаях, когда отображается слишком много кнопок. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **Image.** **Image** — это дочерний элемент **Images**, который в свою очередь является дочерним для элемента **Resources**. Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64.|
|**Tooltip**|Необязательный параметр. Всплывающая подсказка группы. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String.** **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**.|
|**Control**|В каждой группе должен быть по крайней мере один элемент управления. Элемент **управления** может быть **кнопкой или** меню.  Используйте **Меню,** чтобы указать список элементов управления кнопками. В настоящее время поддерживаются только кнопки и меню. Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](control.md#button-control) и [Элементы управления меню](control.md#menu-dropdown-button-controls).<br/>**Примечание:**  Чтобы упростить устранение неполадок, рекомендуется добавлять по одному **элементу Control** и связанным с ними детским элементам **Resources.**|
|**Script**|Ссылка на файл JavaScript с пользовательским определением функции и кодом регистрации. Этот элемент не используется в предварительной версии для разработчиков. Загрузку всех файлов JavaScript выполняет страница HTML.|
|**Page**|Ссылка на HTML-страницу для пользовательских функций.|

## <a name="extension-points-for-outlook"></a>Точки расширения для Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (можно использовать только в [DesktopFormFactor](desktopformfactor.md))
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface)
- [LaunchEvent](#launchevent)
- [Events](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface

Эта точка расширения помещает кнопки на панель команд для представления чтения почты. В классической версии Outlook эта панель отображается на ленте.

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

> [!IMPORTANT]
> Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна в этой точке расширения.

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

### <a name="mobileonlinemeetingcommandsurface"></a>MobileOnlineMeetingCommandSurface

В этой точке расширения в командной поверхности для встречи в мобильном форм-факторе помещается соответствующий режиму. Организатор собрания может создать собрание в Интернете. Впоследствии участник может присоединиться к собранию в Интернете. Дополнительные статьи об этом сценарии см. в статье [Create an mobile add-in Outlook для поставщика веб-собраний.](../../outlook/online-meeting.md)

> [!NOTE]
> Эта точка расширения поддерживается только на Android и iOS с Microsoft 365 подпиской.
>
> Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна в этой точке расширения.

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [Control](control.md) |  Добавляет кнопку на поверхность команды.  |

`ExtensionPoint` элементы этого типа могут иметь только один детский элемент: `Control` элемент.

Элемент, `Control` содержащийся в этой точке расширения, должен иметь `xsi:type` набор атрибутов `MobileButton` .

Изображения `Icon` должны быть в сером масштабе с использованием кода hex или его `#919191` эквивалента в других [цветовых форматах.](https://convertingcolors.com/hex-color-919191.html)

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

### <a name="launchevent"></a>LaunchEvent

Эта точка расширения позволяет надстройку активировать на основе поддерживаемых событий в форм-факторе рабочего стола. Дополнительные информацию об этом сценарии и полном списке поддерживаемых событий [см.](../../outlook/autolaunch.md) в статье Настройка надстройки Outlook для статьи активации на основе событий.

> [!IMPORTANT]
> Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна в этой точке расширения.

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  Список [launchEvent для](launchevent.md) активации на основе событий.  |
| [SourceLocation](sourcelocation.md) |  Расположение исходных файлов JavaScript.  |

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

Эта точка расширения добавляет обработчик для указанного события. Дополнительные сведения об использовании этой точки расширения см. в пункте [On-send для Outlook надстройки.](../../outlook/outlook-on-send-addins.md)

> [!IMPORTANT]
> Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна в этой точке расширения.

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

> [!IMPORTANT]
> Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна в этой точке расширения.

В соответствующем элементе [VersionOverrides](versionoverrides.md) для атрибута `xsi:type` должно быть задано значение `VersionOverridesV1_1`.

> [!NOTE]
> Этот тип элемента доступен в [клиентах Outlook, поддерживающих наборы обязательных требований 1.6 и более поздних версий.](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)

|  Элемент |  Описание  |
|:-----|:-----|
|  [Label](#label) |  Задает метку для надстройки в контекстном окне.  |
|  [SourceLocation](sourcelocation.md) |  Задает URL-адрес контекстного окна.  |
|  [Rule](rule.md) |  Задает одно или несколько правил, определяющих, когда активируется надстройка.  |

#### <a name="label"></a>Label

Обязательный элемент. Метка группы. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)

#### <a name="highlight-requirements"></a>Требования к выделению

Единственный способ, которым пользователь может активировать контекстную надстройку, — взаимодействие с выделенной сущностью. Разработчики могут указывать, какие сущности выделяются, с помощью атрибута `Highlight` элемента `Rule` для типов правил `ItemHasKnownEntity` и `ItemHasRegularExpressionMatch`.

Однако следует учитывать некоторые ограничения. Они гарантируют, что в соответствующих сообщениях и встречах всегда есть выделенная сущность, с помощью которой пользователь может активировать надстройку.

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
