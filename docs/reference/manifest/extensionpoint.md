---
title: Элемент ExtensionPoint в файле манифеста
description: Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8ccc08a9c0d42edf89c904b8809a530239be4c
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855634"
---
# <a name="extensionpoint-element"></a>Элемент ExtensionPoint

 Определяет, где доступны функции надстройки в пользовательском интерфейсе Office. Элемент **ExtensionPoint** является дочерним для элемента [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Да  | Тип определяемой точки расширения. Возможные значения зависят от Office хост-приложения, определенного в значении  элемента host-дедушек.|

## <a name="extension-points-for-excel-onenote-powerpoint-and-word-add-in-commands"></a>Точки расширения для команд Excel, OneNote, PowerPoint и Word

Существует три типа точек расширения, доступных в некоторых или всех этих хостов.

- [PrimaryCommandSurface](#primarycommandsurface) (Допустимо для Word, Excel, PowerPoint и OneNote) — лента в Office.
- [ContextMenu](#contextmenu) (Допустимо для Word, Excel, PowerPoint и OneNote) — меню ярлыка, которое отображается при выборе и удержании (или правой кнопкой мыши) Office пользовательского интерфейса.
- [CustomFunctions](#customfunctions) (Допустимо только для Excel) — настраиваемая функция, написанная в JavaScript для Excel.

См. следующие подсекции для детских элементов и примеры этих типов точек расширения.

### <a name="primarycommandsurface"></a>PrimaryCommandSurface

Основная поверхность команды в Word, Excel, PowerPoint и OneNote является лентой.

#### <a name="child-elements"></a>Дочерние элементы

|Элемент|Описание|
|:-----|:-----|
|[CustomTab] (customtab.md|Обязательный, если требуется добавить пользовательскую вкладку в ленту (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **CustomTab** и **OfficeTab** одновременно. Атрибут **id** является обязательным. |
|[OfficeTab](officetab.md)|Требуется, если требуется расширить вкладку ленты Приложение Office по умолчанию (с помощью **PrimaryCommandSurface**). Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно.|

#### <a name="example"></a>Пример

В следующем примере показано, как использовать **элемент ExtensionPoint** с **PrimaryCommandSurface**. В ленту добавляется настраиваемая вкладка.

> [!IMPORTANT]
> Убедитесь, что для элементов, которые содержат атрибут ID, указан уникальный идентификатор.

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.MyTab1">
    <Label resid="residLabel4" />
    <Group id="Contoso.Group1">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Contoso.Button1">
          <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
```

### <a name="contextmenu"></a>ContextMenu

Меню контекста — это меню ярлыка, которое появляется при щелчке правой кнопкой мыши в Office пользовательского интерфейса.

#### <a name="child-elements"></a>Дочерние элементы
 
|Элемент|Описание|
|:-----|:-----|
|[OfficeMenu](officemenu.md)|Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Атрибут **id** должен быть установлен к одной из следующих строк: <br/> - **ContextMenuText** , если контекстное меню должно открываться, когда пользователь щелкает правой кнопкой мыши по выбранному тексту. <br/> - **ContextMenuCell**, если контекстное меню должно открываться, когда пользователь щелкает правой кнопкой мыши по ячейке на Excel таблице.|

#### <a name="example"></a>Пример

Ниже добавлено настраиваемое контекстное меню для ячеек в Excel таблицы.

```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

### <a name="customfunctions"></a>Настраиваемые функции

Настраиваемая функция, написанная в JavaScript или TypeScript для Excel.

#### <a name="child-elements"></a>Дочерние элементы

|Элемент|Описание|
|:-----|:-----|
|[Script](script.md)|Обязательный элемент. Ссылки на файл JavaScript с определением и кодом регистрации пользовательской функции.|
|[Page](page.md)|Обязательный элемент. Ссылка на HTML-страницу для пользовательских функций.|
|[Метаданные](metadata.md)|Обязательный элемент. Определяет параметры метаданных, используемые пользовательской функцией в Excel.|
|[Namespace](namespace.md)|Необязательное свойство. Определяет пространство имен, используемых пользовательской функцией в Excel.|

#### <a name="example"></a>Пример

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="Functions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="Shared.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
  <Namespace resid="Functions.Namespace"/>
</ExtensionPoint>
```

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
  <CustomTab id="Contoso.TabCustom2">
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
  <CustomTab id="Contoso.TabCustom3">
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
  <CustomTab id="Contoso.TabCustom4">
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
  <CustomTab id="Contoso.TabCustom5">
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
  <Group id="Contoso.mobileGroup1">
    <Label resid="residAppName"/>
      <Control  xsi:type="MobileButton id="Contoso.mobileButton1"">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface"></a>MobileOnlineMeetingCommandSurface

В этой точке расширения в командной поверхности для встречи в мобильном форм-факторе помещается соответствующий режиму. Организатор собрания может создать собрание в Интернете. Впоследствии участник может присоединиться к собранию в Интернете. Дополнительные статьи об этом сценарии см. в статье [Создание надстройки Outlook для поставщика веб-собраний](../../outlook/online-meeting.md).

> [!NOTE]
> Эта точка расширения поддерживается только на Android и iOS с Microsoft 365 подпиской.
>
> Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна в этой точке расширения.

#### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [Control](control.md) |  Добавляет кнопку на поверхность команды.  |

`ExtensionPoint` элементы этого типа могут иметь только один детский элемент: элемент `Control` .

Элемент `Control` , содержащийся в этой точке расширения, должен иметь набор `xsi:type` атрибутов `MobileButton`.

Изображения `Icon` должны быть в сером масштабе с использованием кода hex или `#919191` его эквивалента в [других цветовых форматах](https://convertingcolors.com/hex-color-919191.html).

#### <a name="example"></a>Пример

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="Contoso.onlineMeetingFunctionButton1">
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

Эта точка расширения позволяет надстройку активировать на основе поддерживаемых событий в форм-факторе рабочего стола. Дополнительные информацию об этом сценарии и полном списке поддерживаемых событий см. в статье [Настройка надстройки Outlook](../../outlook/autolaunch.md) для активации на основе событий.

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

Эта точка расширения добавляет обработчик для указанного события. Дополнительные сведения об использовании этой точки расширения см. в пункте [On-send для Outlook надстройки](../../outlook/outlook-on-send-addins.md).

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

Обязательный элемент. Метка группы. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources](resources.md) .

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
