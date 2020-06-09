---
title: Элемент ExtensionPoint в файле манифеста
description: Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 56a7365f986060a225cc0d20c89a310deb4a52b3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611871"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="e9f65-103">Элемент ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="e9f65-103">ExtensionPoint element</span></span>

 <span data-ttu-id="e9f65-104">Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="e9f65-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="e9f65-105">Элемент **ExtensionPoint** является дочерним для элемента [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e9f65-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="e9f65-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e9f65-106">Attributes</span></span>

|  <span data-ttu-id="e9f65-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="e9f65-107">Attribute</span></span>  |  <span data-ttu-id="e9f65-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e9f65-108">Required</span></span>  |  <span data-ttu-id="e9f65-109">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e9f65-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="e9f65-110">**xsi:type**</span></span>  |  <span data-ttu-id="e9f65-111">Да</span><span class="sxs-lookup"><span data-stu-id="e9f65-111">Yes</span></span>  | <span data-ttu-id="e9f65-112">Тип определяемой точки расширения.</span><span class="sxs-lookup"><span data-stu-id="e9f65-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="e9f65-113">Точки расширения только для Excel</span><span class="sxs-lookup"><span data-stu-id="e9f65-113">Extension points for Excel only</span></span>

- <span data-ttu-id="e9f65-114">**CustomFunctions** — пользовательская функция, написанная на JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="e9f65-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="e9f65-115">[В этом примере кода XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) показано, как использовать элемент **ExtensionPoint** со значением атрибута **CustomFunctions** и какие дочерние элементы следует использовать.</span><span class="sxs-lookup"><span data-stu-id="e9f65-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="e9f65-116">Точки расширения для команд надстроек Word, Excel, PowerPoint и OneNote</span><span class="sxs-lookup"><span data-stu-id="e9f65-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="e9f65-117">**PrimaryCommandSurface** — лента в Office.</span><span class="sxs-lookup"><span data-stu-id="e9f65-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="e9f65-118">**ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="e9f65-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="e9f65-119">В приведенных ниже примерах показано, как применять элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.</span><span class="sxs-lookup"><span data-stu-id="e9f65-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e9f65-p102">Для элементов, которые содержат атрибут ID, обязательно предоставляйте уникальный идентификатор. Мы рекомендуем использовать название вашей компании и личный идентификатор. Пример формата приведен ниже. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="e9f65-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="e9f65-123">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-123">Child elements</span></span>
 
|<span data-ttu-id="e9f65-124">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="e9f65-124">**Element**</span></span>|<span data-ttu-id="e9f65-125">**Описание**</span><span class="sxs-lookup"><span data-stu-id="e9f65-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="e9f65-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="e9f65-126">**CustomTab**</span></span>|<span data-ttu-id="e9f65-p103">Обязательный, если требуется добавить пользовательскую вкладку в ленту (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **CustomTab** и **OfficeTab** одновременно. Атрибут **id** является обязательным. </span><span class="sxs-lookup"><span data-stu-id="e9f65-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="e9f65-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="e9f65-130">**OfficeTab**</span></span>|<span data-ttu-id="e9f65-131">Обязательный, если требуется расширить стандартную вкладку ленты Office (с помощью элемента **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="e9f65-131">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="e9f65-132">Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно.</span><span class="sxs-lookup"><span data-stu-id="e9f65-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="e9f65-133">Дополнительные сведения см. в разделе [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="e9f65-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="e9f65-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="e9f65-134">**OfficeMenu**</span></span>|<span data-ttu-id="e9f65-p105">Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: </span><span class="sxs-lookup"><span data-stu-id="e9f65-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="e9f65-p106">- **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши. </span><span class="sxs-lookup"><span data-stu-id="e9f65-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="e9f65-p107">- **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь нажимает ячейку электронной таблицы правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="e9f65-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="e9f65-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="e9f65-141">**Group**</span></span>|<span data-ttu-id="e9f65-p108">Группа точек расширения интерфейса пользователя на вкладке. В группе может быть до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов. </span><span class="sxs-lookup"><span data-stu-id="e9f65-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="e9f65-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="e9f65-145">**Label**</span></span>|<span data-ttu-id="e9f65-p109">Обязательный. Метка группы. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="e9f65-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="e9f65-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="e9f65-150">**Icon**</span></span>|<span data-ttu-id="e9f65-p110">Обязательный. Определяет значок группы для использования на устройствах с малым форм-фактором или в случаях, когда отображается слишком много кнопок. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **Image**. **Image** — это дочерний элемент **Images**, который в свою очередь является дочерним для элемента **Resources**. Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64. </span><span class="sxs-lookup"><span data-stu-id="e9f65-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="e9f65-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="e9f65-158">**Tooltip**</span></span>|<span data-ttu-id="e9f65-p111">Необязательный параметр. Всплывающая подсказка группы. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="e9f65-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="e9f65-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="e9f65-163">**Control**</span></span>|<span data-ttu-id="e9f65-164">В каждой группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="e9f65-164">Each group requires at least one control.</span></span> <span data-ttu-id="e9f65-165">Элемент **управления** может быть либо **кнопкой** , либо **меню**.</span><span class="sxs-lookup"><span data-stu-id="e9f65-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="e9f65-166">Используйте **меню** , чтобы указать раскрывающийся список элементов управления "Кнопка".</span><span class="sxs-lookup"><span data-stu-id="e9f65-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="e9f65-167">В настоящее время поддерживаются только кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="e9f65-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="e9f65-168">Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](control.md#button-control) и [Элементы управления меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="e9f65-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="e9f65-169">**Примечание:**  Чтобы упростить устранение неполадок, рекомендуется добавлять элемент **Control** и соответствующие дочерние элементы **Resources** по одному.</span><span class="sxs-lookup"><span data-stu-id="e9f65-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="e9f65-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="e9f65-170">**Script**</span></span>|<span data-ttu-id="e9f65-171">Ссылка на файл JavaScript с пользовательским определением функции и кодом регистрации.</span><span class="sxs-lookup"><span data-stu-id="e9f65-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="e9f65-172">Этот элемент не используется в предварительной версии для разработчиков.</span><span class="sxs-lookup"><span data-stu-id="e9f65-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="e9f65-173">Загрузку всех файлов JavaScript выполняет страница HTML.</span><span class="sxs-lookup"><span data-stu-id="e9f65-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="e9f65-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="e9f65-174">**Page**</span></span>|<span data-ttu-id="e9f65-175">Ссылка на HTML-страницу для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="e9f65-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="e9f65-176">Точки расширения для Outlook</span><span class="sxs-lookup"><span data-stu-id="e9f65-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="e9f65-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="e9f65-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="e9f65-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="e9f65-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="e9f65-181">[Module](#module) (можно использовать только в [DesktopFormFactor](desktopformfactor.md))</span><span class="sxs-lookup"><span data-stu-id="e9f65-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="e9f65-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="e9f65-183">мобилеонлинемитингкоммандсурфаце</span><span class="sxs-lookup"><span data-stu-id="e9f65-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="e9f65-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="e9f65-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="e9f65-185">Events</span><span class="sxs-lookup"><span data-stu-id="e9f65-185">Events</span></span>](#events)
- [<span data-ttu-id="e9f65-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="e9f65-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="e9f65-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="e9f65-p114">Эта точка расширения помещает кнопки на панель команд для представления чтения почты. В классической версии Outlook эта панель отображается на ленте.</span><span class="sxs-lookup"><span data-stu-id="e9f65-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="e9f65-190">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-190">Child elements</span></span>

|  <span data-ttu-id="e9f65-191">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-191">Element</span></span> |  <span data-ttu-id="e9f65-192">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="e9f65-194">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e9f65-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="e9f65-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="e9f65-196">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="e9f65-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="e9f65-197">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="e9f65-198">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="e9f65-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="e9f65-200">Эта точка расширения добавляет кнопки на ленту для надстроек, использующих форму создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="e9f65-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="e9f65-201">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-201">Child elements</span></span>

|  <span data-ttu-id="e9f65-202">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-202">Element</span></span> |  <span data-ttu-id="e9f65-203">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="e9f65-205">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e9f65-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="e9f65-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="e9f65-207">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="e9f65-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="e9f65-208">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="e9f65-209">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="e9f65-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="e9f65-211">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="e9f65-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="e9f65-212">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-212">Child elements</span></span>

|  <span data-ttu-id="e9f65-213">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-213">Element</span></span> |  <span data-ttu-id="e9f65-214">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="e9f65-216">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e9f65-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="e9f65-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="e9f65-218">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="e9f65-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="e9f65-219">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="e9f65-220">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="e9f65-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="e9f65-222">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для участника собрания.</span><span class="sxs-lookup"><span data-stu-id="e9f65-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="e9f65-223">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-223">Child elements</span></span>

|  <span data-ttu-id="e9f65-224">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-224">Element</span></span> |  <span data-ttu-id="e9f65-225">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="e9f65-227">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e9f65-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="e9f65-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="e9f65-229">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="e9f65-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="e9f65-230">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="e9f65-231">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="e9f65-232">Module</span><span class="sxs-lookup"><span data-stu-id="e9f65-232">Module</span></span>

<span data-ttu-id="e9f65-233">Эта точка расширения добавляет кнопки на ленту для расширения модуля.</span><span class="sxs-lookup"><span data-stu-id="e9f65-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="e9f65-234">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-234">Child elements</span></span>

|  <span data-ttu-id="e9f65-235">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-235">Element</span></span> |  <span data-ttu-id="e9f65-236">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-236">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-237">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-237">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="e9f65-238">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e9f65-238">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="e9f65-239">CustomTab</span><span class="sxs-lookup"><span data-stu-id="e9f65-239">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="e9f65-240">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="e9f65-240">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="e9f65-241">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="e9f65-241">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="e9f65-242">Эта точка расширения помещает кнопки на панель команд для чтения почты в форм-факторе мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="e9f65-242">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="e9f65-243">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-243">Child elements</span></span>

|  <span data-ttu-id="e9f65-244">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-244">Element</span></span> |  <span data-ttu-id="e9f65-245">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-245">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-246">Group</span><span class="sxs-lookup"><span data-stu-id="e9f65-246">Group</span></span>](group.md) |  <span data-ttu-id="e9f65-247">Добавляет группу кнопок на панель команд.</span><span class="sxs-lookup"><span data-stu-id="e9f65-247">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="e9f65-248">У элементов **ExtensionPoint** этого типа может быть только один дочерний элемент **Group**.</span><span class="sxs-lookup"><span data-stu-id="e9f65-248">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="e9f65-249">Для атрибута **xsi:type** элементов **Control**, содержащихся в этой точке расширения, должно быть назначено значение `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="e9f65-249">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="e9f65-250">Пример</span><span class="sxs-lookup"><span data-stu-id="e9f65-250">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="e9f65-251">Мобилеонлинемитингкоммандсурфаце (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="e9f65-251">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="e9f65-252">Эта точка расширения поддерживается только в [предварительной версии](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Android с подпиской на Office 365.</span><span class="sxs-lookup"><span data-stu-id="e9f65-252">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="e9f65-253">Эта точка расширения помещает переключатель, подходящий для режима, на поверхности команды для встречи в мобильном конструктивном параметре.</span><span class="sxs-lookup"><span data-stu-id="e9f65-253">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="e9f65-254">Организатор собрания может создать собрание по сети.</span><span class="sxs-lookup"><span data-stu-id="e9f65-254">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="e9f65-255">Затем участник может присоединиться к собранию по сети.</span><span class="sxs-lookup"><span data-stu-id="e9f65-255">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="e9f65-256">Чтобы узнать больше об этом сценарии, ознакомьтесь со статьей [Создание надстройки Outlook Mobile для веб-службы "поставщик собраний](../../outlook/online-meeting.md) ".</span><span class="sxs-lookup"><span data-stu-id="e9f65-256">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="e9f65-257">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-257">Child elements</span></span>

|  <span data-ttu-id="e9f65-258">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-258">Element</span></span> |  <span data-ttu-id="e9f65-259">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-259">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-260">Control</span><span class="sxs-lookup"><span data-stu-id="e9f65-260">Control</span></span>](control.md) |  <span data-ttu-id="e9f65-261">Добавляет кнопку на поверхность команды.</span><span class="sxs-lookup"><span data-stu-id="e9f65-261">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="e9f65-262">`ExtensionPoint`у элементов этого типа может быть только один дочерний элемент: `Control` element.</span><span class="sxs-lookup"><span data-stu-id="e9f65-262">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="e9f65-263">`Control`Атрибуту элемента, содержащегося в этой точке расширения, должен быть `xsi:type` присвое значение `MobileButton` .</span><span class="sxs-lookup"><span data-stu-id="e9f65-263">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="e9f65-264">`Icon`Изображения должны быть в градациях серого с использованием шестнадцатеричного кода `#919191` или его эквивалента в [других цветовых форматах](https://convertingcolors.com/hex-color-919191.html).</span><span class="sxs-lookup"><span data-stu-id="e9f65-264">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="e9f65-265">Пример</span><span class="sxs-lookup"><span data-stu-id="e9f65-265">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="e9f65-266">Лаунчевент (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="e9f65-266">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="e9f65-267">Эта точка расширения поддерживается только в [предварительном просмотре](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) в Outlook в Интернете с подпиской на Office 365.</span><span class="sxs-lookup"><span data-stu-id="e9f65-267">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with an Office 365 subscription.</span></span>

<span data-ttu-id="e9f65-268">Эта точка расширения позволяет активировать надстройку на основе поддерживаемых событий на настольных формах.</span><span class="sxs-lookup"><span data-stu-id="e9f65-268">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="e9f65-269">В настоящее время единственными поддерживаемыми событиями являются `OnNewMessageCompose` и `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="e9f65-269">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="e9f65-270">Чтобы узнать больше об этом сценарии, ознакомьтесь со статьей [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="e9f65-270">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="e9f65-271">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e9f65-271">Child elements</span></span>

|  <span data-ttu-id="e9f65-272">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-272">Element</span></span> |  <span data-ttu-id="e9f65-273">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-273">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="e9f65-274">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="e9f65-274">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="e9f65-275">Список [лаунчевент](launchevent.md) для активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="e9f65-275">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="e9f65-276">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e9f65-276">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="e9f65-277">Расположение исходного файла JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e9f65-277">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="e9f65-278">Пример</span><span class="sxs-lookup"><span data-stu-id="e9f65-278">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="e9f65-279">События</span><span class="sxs-lookup"><span data-stu-id="e9f65-279">Events</span></span>

<span data-ttu-id="e9f65-280">Эта точка расширения добавляет обработчик для указанного события.</span><span class="sxs-lookup"><span data-stu-id="e9f65-280">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="e9f65-281">Для получения дополнительных сведений об использовании этой точки расширения, ознакомьтесь со статьей [функция On Send для надстроек Outlook](../../outlook/outlook-on-send-addins.md).</span><span class="sxs-lookup"><span data-stu-id="e9f65-281">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

| <span data-ttu-id="e9f65-282">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-282">Element</span></span> | <span data-ttu-id="e9f65-283">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-283">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-284">Event</span><span class="sxs-lookup"><span data-stu-id="e9f65-284">Event</span></span>](event.md) |  <span data-ttu-id="e9f65-285">Задает событие и функцию его обработчика.</span><span class="sxs-lookup"><span data-stu-id="e9f65-285">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="e9f65-286">Пример события ItemSend</span><span class="sxs-lookup"><span data-stu-id="e9f65-286">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="e9f65-287">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="e9f65-287">DetectedEntity</span></span>

<span data-ttu-id="e9f65-288">Эта точка расширения добавляет активацию контекстной надстройки для указанного типа сущности.</span><span class="sxs-lookup"><span data-stu-id="e9f65-288">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="e9f65-289">В соответствующем элементе [VersionOverrides](versionoverrides.md) для атрибута `xsi:type` должно быть задано значение `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="e9f65-289">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="e9f65-290">Этот тип элемента доступен в [клиентах Outlook, поддерживающих наборы обязательных требований 1.6 и более поздних версий.](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)</span><span class="sxs-lookup"><span data-stu-id="e9f65-290">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="e9f65-291">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9f65-291">Element</span></span> |  <span data-ttu-id="e9f65-292">Описание</span><span class="sxs-lookup"><span data-stu-id="e9f65-292">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="e9f65-293">Label</span><span class="sxs-lookup"><span data-stu-id="e9f65-293">Label</span></span>](#label) |  <span data-ttu-id="e9f65-294">Задает метку для надстройки в контекстном окне.</span><span class="sxs-lookup"><span data-stu-id="e9f65-294">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="e9f65-295">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e9f65-295">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="e9f65-296">Задает URL-адрес контекстного окна.</span><span class="sxs-lookup"><span data-stu-id="e9f65-296">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="e9f65-297">Rule</span><span class="sxs-lookup"><span data-stu-id="e9f65-297">Rule</span></span>](rule.md) |  <span data-ttu-id="e9f65-298">Задает одно или несколько правил, определяющих, когда активируется надстройка.</span><span class="sxs-lookup"><span data-stu-id="e9f65-298">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="e9f65-299">Label</span><span class="sxs-lookup"><span data-stu-id="e9f65-299">Label</span></span>

<span data-ttu-id="e9f65-300">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="e9f65-300">Required.</span></span> <span data-ttu-id="e9f65-301">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="e9f65-301">The label of the group.</span></span> <span data-ttu-id="e9f65-302">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="e9f65-302">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="e9f65-303">Требования к выделению</span><span class="sxs-lookup"><span data-stu-id="e9f65-303">Highlight requirements</span></span>

<span data-ttu-id="e9f65-p119">Единственный способ, которым пользователь может активировать контекстную надстройку, — взаимодействие с выделенной сущностью. Разработчики могут указывать, какие сущности выделяются, с помощью атрибута `Highlight` элемента `Rule` для типов правил `ItemHasKnownEntity` и `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="e9f65-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="e9f65-p120">Однако следует учитывать некоторые ограничения. Они гарантируют, что в соответствующих сообщениях и встречах всегда есть выделенная сущность, с помощью которой пользователь может активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="e9f65-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="e9f65-308">Сущности `EmailAddress` и `Url` не поддерживают выделение, поэтому их нельзя использовать для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="e9f65-308">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="e9f65-309">Если используется одно правило, то для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="e9f65-309">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="e9f65-310">Если используется правило `RuleCollection`, совмещенное с другими правилами с помощью оператора `Mode="AND"`, то как минимум в одном из правил для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="e9f65-310">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="e9f65-311">Если используется правило `RuleCollection`, в котором правила совмещаются с помощью оператора `Mode="OR"`, то в каждом из них для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="e9f65-311">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="e9f65-312">Пример события DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="e9f65-312">DetectedEntity event example</span></span>

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
