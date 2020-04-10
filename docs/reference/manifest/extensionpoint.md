---
title: Элемент ExtensionPoint в файле манифеста
description: Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 89040ff84afd7e4c33bea8af04255ef19a8e0f50
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215091"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="6763e-103">Элемент ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="6763e-103">ExtensionPoint element</span></span>

 <span data-ttu-id="6763e-104">Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="6763e-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="6763e-105">Элемент **ExtensionPoint** является дочерним для элемента [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="6763e-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="6763e-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6763e-106">Attributes</span></span>

|  <span data-ttu-id="6763e-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="6763e-107">Attribute</span></span>  |  <span data-ttu-id="6763e-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="6763e-108">Required</span></span>  |  <span data-ttu-id="6763e-109">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6763e-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="6763e-110">**xsi:type**</span></span>  |  <span data-ttu-id="6763e-111">Да</span><span class="sxs-lookup"><span data-stu-id="6763e-111">Yes</span></span>  | <span data-ttu-id="6763e-112">Тип определяемой точки расширения.</span><span class="sxs-lookup"><span data-stu-id="6763e-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="6763e-113">Точки расширения только для Excel</span><span class="sxs-lookup"><span data-stu-id="6763e-113">Extension points for Excel only</span></span>

- <span data-ttu-id="6763e-114">**CustomFunctions** — пользовательская функция, написанная на JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="6763e-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="6763e-115">[В этом примере кода XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) показано, как использовать элемент **ExtensionPoint** со значением атрибута **CustomFunctions** и какие дочерние элементы следует использовать.</span><span class="sxs-lookup"><span data-stu-id="6763e-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="6763e-116">Точки расширения для команд надстроек Word, Excel, PowerPoint и OneNote</span><span class="sxs-lookup"><span data-stu-id="6763e-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="6763e-117">**PrimaryCommandSurface** — лента в Office.</span><span class="sxs-lookup"><span data-stu-id="6763e-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="6763e-118">**ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="6763e-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="6763e-119">В приведенных ниже примерах показано, как применять элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.</span><span class="sxs-lookup"><span data-stu-id="6763e-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6763e-p102">Для элементов, которые содержат атрибут ID, обязательно предоставляйте уникальный идентификатор. Мы рекомендуем использовать название вашей компании и личный идентификатор. Пример формата приведен ниже. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="6763e-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="6763e-123">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6763e-123">Child elements</span></span>
 
|<span data-ttu-id="6763e-124">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="6763e-124">**Element**</span></span>|<span data-ttu-id="6763e-125">**Описание**</span><span class="sxs-lookup"><span data-stu-id="6763e-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="6763e-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="6763e-126">**CustomTab**</span></span>|<span data-ttu-id="6763e-p103">Обязательный, если требуется добавить пользовательскую вкладку в ленту (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **CustomTab** и **OfficeTab** одновременно. Атрибут **id** является обязательным. </span><span class="sxs-lookup"><span data-stu-id="6763e-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="6763e-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="6763e-130">**OfficeTab**</span></span>|<span data-ttu-id="6763e-131">Обязательный, если требуется расширить стандартную вкладку ленты Office (с помощью элемента **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="6763e-131">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="6763e-132">Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно.</span><span class="sxs-lookup"><span data-stu-id="6763e-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="6763e-133">Дополнительные сведения см. в разделе [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="6763e-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="6763e-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="6763e-134">**OfficeMenu**</span></span>|<span data-ttu-id="6763e-p105">Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: </span><span class="sxs-lookup"><span data-stu-id="6763e-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="6763e-p106">- **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши. </span><span class="sxs-lookup"><span data-stu-id="6763e-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="6763e-p107">- **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь нажимает ячейку электронной таблицы правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="6763e-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="6763e-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="6763e-141">**Group**</span></span>|<span data-ttu-id="6763e-p108">Группа точек расширения интерфейса пользователя на вкладке. В группе может быть до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов. </span><span class="sxs-lookup"><span data-stu-id="6763e-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="6763e-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="6763e-145">**Label**</span></span>|<span data-ttu-id="6763e-p109">Обязательный. Метка группы. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="6763e-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="6763e-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="6763e-150">**Icon**</span></span>|<span data-ttu-id="6763e-p110">Обязательный. Определяет значок группы для использования на устройствах с малым форм-фактором или в случаях, когда отображается слишком много кнопок. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **Image**. **Image** — это дочерний элемент **Images**, который в свою очередь является дочерним для элемента **Resources**. Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64. </span><span class="sxs-lookup"><span data-stu-id="6763e-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="6763e-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="6763e-158">**Tooltip**</span></span>|<span data-ttu-id="6763e-p111">Необязательный параметр. Всплывающая подсказка группы. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="6763e-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="6763e-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="6763e-163">**Control**</span></span>|<span data-ttu-id="6763e-164">В каждой группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="6763e-164">Each group requires at least one control.</span></span> <span data-ttu-id="6763e-165">Элемент **управления** может быть либо **кнопкой** , либо **меню**.</span><span class="sxs-lookup"><span data-stu-id="6763e-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="6763e-166">Используйте **меню** , чтобы указать раскрывающийся список элементов управления "Кнопка".</span><span class="sxs-lookup"><span data-stu-id="6763e-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="6763e-167">В настоящее время поддерживаются только кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="6763e-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="6763e-168">Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](control.md#button-control) и [Элементы управления меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="6763e-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="6763e-169">**Примечание:**  Чтобы упростить устранение неполадок, рекомендуется добавлять элемент **Control** и соответствующие дочерние элементы **Resources** по одному.</span><span class="sxs-lookup"><span data-stu-id="6763e-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="6763e-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="6763e-170">**Script**</span></span>|<span data-ttu-id="6763e-171">Ссылка на файл JavaScript с пользовательским определением функции и кодом регистрации.</span><span class="sxs-lookup"><span data-stu-id="6763e-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="6763e-172">Этот элемент не используется в предварительной версии для разработчиков.</span><span class="sxs-lookup"><span data-stu-id="6763e-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="6763e-173">Загрузку всех файлов JavaScript выполняет страница HTML.</span><span class="sxs-lookup"><span data-stu-id="6763e-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="6763e-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="6763e-174">**Page**</span></span>|<span data-ttu-id="6763e-175">Ссылка на HTML-страницу для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="6763e-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="6763e-176">Точки расширения для Outlook</span><span class="sxs-lookup"><span data-stu-id="6763e-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="6763e-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="6763e-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="6763e-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="6763e-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="6763e-181">[Module](#module) (можно использовать только в [DesktopFormFactor](desktopformfactor.md))</span><span class="sxs-lookup"><span data-stu-id="6763e-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="6763e-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="6763e-183">мобилеонлинемитингкоммандсурфаце</span><span class="sxs-lookup"><span data-stu-id="6763e-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="6763e-184">Events</span><span class="sxs-lookup"><span data-stu-id="6763e-184">Events</span></span>](#events)
- [<span data-ttu-id="6763e-185">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="6763e-185">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="6763e-186">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-186">MessageReadCommandSurface</span></span>

<span data-ttu-id="6763e-p114">Эта точка расширения помещает кнопки на панель команд для представления чтения почты. В классической версии Outlook эта панель отображается на ленте.</span><span class="sxs-lookup"><span data-stu-id="6763e-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="6763e-189">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6763e-189">Child elements</span></span>

|  <span data-ttu-id="6763e-190">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-190">Element</span></span> |  <span data-ttu-id="6763e-191">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-191">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-192">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-192">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="6763e-193">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="6763e-193">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="6763e-194">CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-194">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="6763e-195">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="6763e-195">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="6763e-196">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-196">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="6763e-197">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-197">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="6763e-198">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-198">MessageComposeCommandSurface</span></span>

<span data-ttu-id="6763e-199">Эта точка расширения добавляет кнопки на ленту для надстроек, использующих форму создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="6763e-199">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="6763e-200">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6763e-200">Child elements</span></span>

|  <span data-ttu-id="6763e-201">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-201">Element</span></span> |  <span data-ttu-id="6763e-202">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-202">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-203">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-203">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="6763e-204">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="6763e-204">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="6763e-205">CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-205">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="6763e-206">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="6763e-206">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="6763e-207">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-207">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="6763e-208">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-208">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="6763e-209">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-209">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="6763e-210">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="6763e-210">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="6763e-211">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6763e-211">Child elements</span></span>

|  <span data-ttu-id="6763e-212">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-212">Element</span></span> |  <span data-ttu-id="6763e-213">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-213">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-214">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-214">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="6763e-215">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="6763e-215">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="6763e-216">CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-216">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="6763e-217">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="6763e-217">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="6763e-218">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-218">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="6763e-219">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-219">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="6763e-220">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-220">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="6763e-221">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для участника собрания.</span><span class="sxs-lookup"><span data-stu-id="6763e-221">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="6763e-222">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6763e-222">Child elements</span></span>

|  <span data-ttu-id="6763e-223">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-223">Element</span></span> |  <span data-ttu-id="6763e-224">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-224">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-225">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-225">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="6763e-226">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="6763e-226">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="6763e-227">CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-227">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="6763e-228">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="6763e-228">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="6763e-229">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-229">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="6763e-230">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-230">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="6763e-231">Module</span><span class="sxs-lookup"><span data-stu-id="6763e-231">Module</span></span>

<span data-ttu-id="6763e-232">Эта точка расширения добавляет кнопки на ленту для расширения модуля.</span><span class="sxs-lookup"><span data-stu-id="6763e-232">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="6763e-233">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6763e-233">Child elements</span></span>

|  <span data-ttu-id="6763e-234">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-234">Element</span></span> |  <span data-ttu-id="6763e-235">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-235">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-236">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="6763e-236">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="6763e-237">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="6763e-237">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="6763e-238">CustomTab</span><span class="sxs-lookup"><span data-stu-id="6763e-238">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="6763e-239">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="6763e-239">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="6763e-240">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="6763e-240">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="6763e-241">Эта точка расширения помещает кнопки на панель команд для чтения почты в форм-факторе мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="6763e-241">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="6763e-242">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6763e-242">Child elements</span></span>

|  <span data-ttu-id="6763e-243">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-243">Element</span></span> |  <span data-ttu-id="6763e-244">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-244">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-245">Group</span><span class="sxs-lookup"><span data-stu-id="6763e-245">Group</span></span>](group.md) |  <span data-ttu-id="6763e-246">Добавляет группу кнопок на панель команд.</span><span class="sxs-lookup"><span data-stu-id="6763e-246">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="6763e-247">У элементов **ExtensionPoint** этого типа может быть только один дочерний элемент **Group**.</span><span class="sxs-lookup"><span data-stu-id="6763e-247">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="6763e-248">Для атрибута **xsi:type** элементов **Control**, содержащихся в этой точке расширения, должно быть назначено значение `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="6763e-248">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="6763e-249">Пример</span><span class="sxs-lookup"><span data-stu-id="6763e-249">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="6763e-250">Мобилеонлинемитингкоммандсурфаце (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="6763e-250">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="6763e-251">Эта точка расширения поддерживается только в [предварительной версии](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Android с подпиской на Office 365.</span><span class="sxs-lookup"><span data-stu-id="6763e-251">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="6763e-252">Эта точка расширения помещает переключатель, подходящий для режима, на поверхности команды для встречи в мобильном конструктивном параметре.</span><span class="sxs-lookup"><span data-stu-id="6763e-252">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="6763e-253">Организатор собрания может создать собрание по сети.</span><span class="sxs-lookup"><span data-stu-id="6763e-253">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="6763e-254">Затем участник может присоединиться к собранию по сети.</span><span class="sxs-lookup"><span data-stu-id="6763e-254">An attendee can subsequently join the online meeting.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="6763e-255">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6763e-255">Child elements</span></span>

|  <span data-ttu-id="6763e-256">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-256">Element</span></span> |  <span data-ttu-id="6763e-257">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-257">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-258">Control</span><span class="sxs-lookup"><span data-stu-id="6763e-258">Control</span></span>](control.md) |  <span data-ttu-id="6763e-259">Добавляет кнопку на поверхность команды.</span><span class="sxs-lookup"><span data-stu-id="6763e-259">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="6763e-260">У элементов **ExtensionPoint** этого типа может быть только один дочерний элемент: элемент **Control** .</span><span class="sxs-lookup"><span data-stu-id="6763e-260">**ExtensionPoint** elements of this type can only have one child element: a **Control** element.</span></span>

<span data-ttu-id="6763e-261">Элемент **Control** , содержащийся в этой точке расширения, должен иметь атрибут **xsi: Type** со `MobileButton`значением.</span><span class="sxs-lookup"><span data-stu-id="6763e-261">The **Control** element contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="6763e-262">Пример</span><span class="sxs-lookup"><span data-stu-id="6763e-262">Example</span></span>

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
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>UiLessInsertMeetingLink</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="6763e-263">События</span><span class="sxs-lookup"><span data-stu-id="6763e-263">Events</span></span>

<span data-ttu-id="6763e-264">Эта точка расширения добавляет обработчик для указанного события.</span><span class="sxs-lookup"><span data-stu-id="6763e-264">This extension point adds an event handler for a specified event.</span></span>

| <span data-ttu-id="6763e-265">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-265">Element</span></span> | <span data-ttu-id="6763e-266">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-266">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-267">Event</span><span class="sxs-lookup"><span data-stu-id="6763e-267">Event</span></span>](event.md) |  <span data-ttu-id="6763e-268">Задает событие и функцию его обработчика.</span><span class="sxs-lookup"><span data-stu-id="6763e-268">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="6763e-269">Пример события ItemSend</span><span class="sxs-lookup"><span data-stu-id="6763e-269">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="6763e-270">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="6763e-270">DetectedEntity</span></span>

<span data-ttu-id="6763e-271">Эта точка расширения добавляет активацию контекстной надстройки для указанного типа сущности.</span><span class="sxs-lookup"><span data-stu-id="6763e-271">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="6763e-272">В соответствующем элементе [VersionOverrides](versionoverrides.md) для атрибута `xsi:type` должно быть задано значение `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="6763e-272">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="6763e-273">Этот тип элемента доступен в [клиентах Outlook, поддерживающих наборы обязательных требований 1.6 и более поздних версий.](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)</span><span class="sxs-lookup"><span data-stu-id="6763e-273">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="6763e-274">Элемент</span><span class="sxs-lookup"><span data-stu-id="6763e-274">Element</span></span> |  <span data-ttu-id="6763e-275">Описание</span><span class="sxs-lookup"><span data-stu-id="6763e-275">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="6763e-276">Label</span><span class="sxs-lookup"><span data-stu-id="6763e-276">Label</span></span>](#label) |  <span data-ttu-id="6763e-277">Задает метку для надстройки в контекстном окне.</span><span class="sxs-lookup"><span data-stu-id="6763e-277">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="6763e-278">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="6763e-278">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="6763e-279">Задает URL-адрес контекстного окна.</span><span class="sxs-lookup"><span data-stu-id="6763e-279">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="6763e-280">Rule</span><span class="sxs-lookup"><span data-stu-id="6763e-280">Rule</span></span>](rule.md) |  <span data-ttu-id="6763e-281">Задает одно или несколько правил, определяющих, когда активируется надстройка.</span><span class="sxs-lookup"><span data-stu-id="6763e-281">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="6763e-282">Label</span><span class="sxs-lookup"><span data-stu-id="6763e-282">Label</span></span>

<span data-ttu-id="6763e-283">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="6763e-283">Required.</span></span> <span data-ttu-id="6763e-284">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="6763e-284">The label of the group.</span></span> <span data-ttu-id="6763e-285">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="6763e-285">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="6763e-286">Требования к выделению</span><span class="sxs-lookup"><span data-stu-id="6763e-286">Highlight requirements</span></span>

<span data-ttu-id="6763e-p117">Единственный способ, которым пользователь может активировать контекстную надстройку, — взаимодействие с выделенной сущностью. Разработчики могут указывать, какие сущности выделяются, с помощью атрибута `Highlight` элемента `Rule` для типов правил `ItemHasKnownEntity` и `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="6763e-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="6763e-p118">Однако следует учитывать некоторые ограничения. Они гарантируют, что в соответствующих сообщениях и встречах всегда есть выделенная сущность, с помощью которой пользователь может активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="6763e-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="6763e-291">Сущности `EmailAddress` и `Url` не поддерживают выделение, поэтому их нельзя использовать для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="6763e-291">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="6763e-292">Если используется одно правило, то для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="6763e-292">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="6763e-293">Если используется правило `RuleCollection`, совмещенное с другими правилами с помощью оператора `Mode="AND"`, то как минимум в одном из правил для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="6763e-293">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="6763e-294">Если используется правило `RuleCollection`, в котором правила совмещаются с помощью оператора `Mode="OR"`, то в каждом из них для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="6763e-294">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="6763e-295">Пример события DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="6763e-295">DetectedEntity event example</span></span>

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
