---
title: Элемент ExtensionPoint в файле манифеста
description: ''
ms.date: 09/05/2019
localization_priority: Normal
ms.openlocfilehash: 44075bd12c15b4ac9117a51d71fdcc7d6436a7ce
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324878"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="615c9-102">Элемент ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="615c9-102">ExtensionPoint element</span></span>

 <span data-ttu-id="615c9-103">Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="615c9-103">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="615c9-104">Элемент **ExtensionPoint** является дочерним для элемента [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="615c9-104">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="615c9-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="615c9-105">Attributes</span></span>

|  <span data-ttu-id="615c9-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="615c9-106">Attribute</span></span>  |  <span data-ttu-id="615c9-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="615c9-107">Required</span></span>  |  <span data-ttu-id="615c9-108">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="615c9-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="615c9-109">**xsi:type**</span></span>  |  <span data-ttu-id="615c9-110">Да</span><span class="sxs-lookup"><span data-stu-id="615c9-110">Yes</span></span>  | <span data-ttu-id="615c9-111">Тип определяемой точки расширения.</span><span class="sxs-lookup"><span data-stu-id="615c9-111">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="615c9-112">Точки расширения только для Excel</span><span class="sxs-lookup"><span data-stu-id="615c9-112">Extension points for Excel only</span></span>

- <span data-ttu-id="615c9-113">**CustomFunctions** — пользовательская функция, написанная на JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="615c9-113">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="615c9-114">[В этом примере кода XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) показано, как использовать элемент **ExtensionPoint** со значением атрибута **CustomFunctions** и какие дочерние элементы следует использовать.</span><span class="sxs-lookup"><span data-stu-id="615c9-114">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="615c9-115">Точки расширения для команд надстроек Word, Excel, PowerPoint и OneNote</span><span class="sxs-lookup"><span data-stu-id="615c9-115">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="615c9-116">**PrimaryCommandSurface** — лента в Office.</span><span class="sxs-lookup"><span data-stu-id="615c9-116">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="615c9-117">**ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="615c9-117">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="615c9-118">В приведенных ниже примерах показано, как применять элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.</span><span class="sxs-lookup"><span data-stu-id="615c9-118">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="615c9-p102">Для элементов, которые содержат атрибут ID, обязательно предоставляйте уникальный идентификатор. Мы рекомендуем использовать название вашей компании и личный идентификатор. Пример формата приведен ниже. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="615c9-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="615c9-122">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="615c9-122">Child elements</span></span>
 
|<span data-ttu-id="615c9-123">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="615c9-123">**Element**</span></span>|<span data-ttu-id="615c9-124">**Описание**</span><span class="sxs-lookup"><span data-stu-id="615c9-124">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="615c9-125">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="615c9-125">**CustomTab**</span></span>|<span data-ttu-id="615c9-p103">Обязательный, если требуется добавить пользовательскую вкладку в ленту (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **CustomTab** и **OfficeTab** одновременно. Атрибут **id** является обязательным. </span><span class="sxs-lookup"><span data-stu-id="615c9-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="615c9-129">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="615c9-129">**OfficeTab**</span></span>|<span data-ttu-id="615c9-p104">Является обязательным, если вы хотите расширить вкладку ленты Office по умолчанию (с помощью **PrimaryCommandSurface**). При использовании элемента **OfficeTab** нельзя использовать элемент **CustomTab** . Дополнительные сведения см. в разделе [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="615c9-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="615c9-133">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="615c9-133">**OfficeMenu**</span></span>|<span data-ttu-id="615c9-p105">Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: </span><span class="sxs-lookup"><span data-stu-id="615c9-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="615c9-p106">- **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши. </span><span class="sxs-lookup"><span data-stu-id="615c9-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="615c9-p107">- **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь нажимает ячейку электронной таблицы правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="615c9-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="615c9-140">**Group**</span><span class="sxs-lookup"><span data-stu-id="615c9-140">**Group**</span></span>|<span data-ttu-id="615c9-p108">Группа точек расширения интерфейса пользователя на вкладке. В группе может быть до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов. </span><span class="sxs-lookup"><span data-stu-id="615c9-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="615c9-144">**Label**</span><span class="sxs-lookup"><span data-stu-id="615c9-144">**Label**</span></span>|<span data-ttu-id="615c9-p109">Обязательный. Метка группы. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="615c9-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="615c9-149">**Icon**</span><span class="sxs-lookup"><span data-stu-id="615c9-149">**Icon**</span></span>|<span data-ttu-id="615c9-p110">Обязательный. Определяет значок группы для использования на устройствах с малым форм-фактором или в случаях, когда отображается слишком много кнопок. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **Image**. **Image** — это дочерний элемент **Images**, который в свою очередь является дочерним для элемента **Resources**. Атрибут **size** определяет размер изображения в пикселях. Обязательными являются три размера изображения: 16, 32 и 80. Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64. </span><span class="sxs-lookup"><span data-stu-id="615c9-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="615c9-157">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="615c9-157">**Tooltip**</span></span>|<span data-ttu-id="615c9-p111">Необязательный параметр. Всплывающая подсказка группы. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**. </span><span class="sxs-lookup"><span data-stu-id="615c9-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="615c9-162">**Control**</span><span class="sxs-lookup"><span data-stu-id="615c9-162">**Control**</span></span>|<span data-ttu-id="615c9-163">В каждой группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="615c9-163">Each group requires at least one control.</span></span> <span data-ttu-id="615c9-164">Элемент **управления** может быть либо **кнопкой** , либо **меню**.</span><span class="sxs-lookup"><span data-stu-id="615c9-164">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="615c9-165">Используйте **меню** , чтобы указать раскрывающийся список элементов управления "Кнопка".</span><span class="sxs-lookup"><span data-stu-id="615c9-165">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="615c9-166">В настоящее время поддерживаются только кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="615c9-166">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="615c9-167">Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](control.md#button-control) и [Элементы управления меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="615c9-167">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="615c9-168">**Примечание:**  Чтобы упростить устранение неполадок, рекомендуется добавлять элемент **Control** и соответствующие дочерние элементы **Resources** по одному.</span><span class="sxs-lookup"><span data-stu-id="615c9-168">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="615c9-169">**Script**</span><span class="sxs-lookup"><span data-stu-id="615c9-169">**Script**</span></span>|<span data-ttu-id="615c9-170">Ссылка на файл JavaScript с пользовательским определением функции и кодом регистрации.</span><span class="sxs-lookup"><span data-stu-id="615c9-170">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="615c9-171">Этот элемент не используется в предварительной версии для разработчиков.</span><span class="sxs-lookup"><span data-stu-id="615c9-171">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="615c9-172">Загрузку всех файлов JavaScript выполняет страница HTML.</span><span class="sxs-lookup"><span data-stu-id="615c9-172">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="615c9-173">**Page**</span><span class="sxs-lookup"><span data-stu-id="615c9-173">**Page**</span></span>|<span data-ttu-id="615c9-174">Ссылка на HTML-страницу для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="615c9-174">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="615c9-175">Точки расширения для Outlook</span><span class="sxs-lookup"><span data-stu-id="615c9-175">Extension points for Outlook</span></span>

- [<span data-ttu-id="615c9-176">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-176">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="615c9-177">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-177">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="615c9-178">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-178">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="615c9-179">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-179">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="615c9-180">[Module](#module) (можно использовать только в [DesktopFormFactor](desktopformfactor.md))</span><span class="sxs-lookup"><span data-stu-id="615c9-180">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="615c9-181">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-181">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="615c9-182">Events</span><span class="sxs-lookup"><span data-stu-id="615c9-182">Events</span></span>](#events)
- [<span data-ttu-id="615c9-183">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="615c9-183">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="615c9-184">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-184">MessageReadCommandSurface</span></span>
<span data-ttu-id="615c9-p114">Эта точка расширения помещает кнопки на панель команд для представления чтения почты. В классической версии Outlook эта панель отображается на ленте.</span><span class="sxs-lookup"><span data-stu-id="615c9-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="615c9-187">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="615c9-187">Child elements</span></span>

|  <span data-ttu-id="615c9-188">Элемент</span><span class="sxs-lookup"><span data-stu-id="615c9-188">Element</span></span> |  <span data-ttu-id="615c9-189">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-189">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="615c9-190">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-190">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="615c9-191">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="615c9-191">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="615c9-192">CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-192">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="615c9-193">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="615c9-193">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="615c9-194">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-194">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="615c9-195">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-195">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="615c9-196">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-196">MessageComposeCommandSurface</span></span>
<span data-ttu-id="615c9-197">Эта точка расширения добавляет кнопки на ленту для надстроек, использующих форму создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="615c9-197">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="615c9-198">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="615c9-198">Child elements</span></span>

|  <span data-ttu-id="615c9-199">Элемент</span><span class="sxs-lookup"><span data-stu-id="615c9-199">Element</span></span> |  <span data-ttu-id="615c9-200">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-200">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="615c9-201">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-201">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="615c9-202">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="615c9-202">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="615c9-203">CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-203">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="615c9-204">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="615c9-204">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="615c9-205">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-205">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="615c9-206">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-206">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="615c9-207">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-207">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="615c9-208">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="615c9-208">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="615c9-209">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="615c9-209">Child elements</span></span>

|  <span data-ttu-id="615c9-210">Элемент</span><span class="sxs-lookup"><span data-stu-id="615c9-210">Element</span></span> |  <span data-ttu-id="615c9-211">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-211">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="615c9-212">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-212">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="615c9-213">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="615c9-213">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="615c9-214">CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-214">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="615c9-215">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="615c9-215">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="615c9-216">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-216">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="615c9-217">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-217">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="615c9-218">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-218">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="615c9-219">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для участника собрания.</span><span class="sxs-lookup"><span data-stu-id="615c9-219">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="615c9-220">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="615c9-220">Child elements</span></span>

|  <span data-ttu-id="615c9-221">Элемент</span><span class="sxs-lookup"><span data-stu-id="615c9-221">Element</span></span> |  <span data-ttu-id="615c9-222">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-222">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="615c9-223">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-223">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="615c9-224">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="615c9-224">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="615c9-225">CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-225">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="615c9-226">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="615c9-226">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="615c9-227">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-227">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="615c9-228">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-228">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="615c9-229">Module</span><span class="sxs-lookup"><span data-stu-id="615c9-229">Module</span></span>

<span data-ttu-id="615c9-230">Эта точка расширения добавляет кнопки на ленту для расширения модуля.</span><span class="sxs-lookup"><span data-stu-id="615c9-230">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="615c9-231">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="615c9-231">Child elements</span></span>

|  <span data-ttu-id="615c9-232">Элемент</span><span class="sxs-lookup"><span data-stu-id="615c9-232">Element</span></span> |  <span data-ttu-id="615c9-233">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-233">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="615c9-234">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="615c9-234">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="615c9-235">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="615c9-235">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="615c9-236">CustomTab</span><span class="sxs-lookup"><span data-stu-id="615c9-236">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="615c9-237">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="615c9-237">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="615c9-238">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="615c9-238">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="615c9-239">Эта точка расширения помещает кнопки на панель команд для чтения почты в форм-факторе мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="615c9-239">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="615c9-240">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="615c9-240">Child elements</span></span>

|  <span data-ttu-id="615c9-241">Элемент</span><span class="sxs-lookup"><span data-stu-id="615c9-241">Element</span></span> |  <span data-ttu-id="615c9-242">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-242">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="615c9-243">Group</span><span class="sxs-lookup"><span data-stu-id="615c9-243">Group</span></span>](group.md) |  <span data-ttu-id="615c9-244">Добавляет группу кнопок на панель команд.</span><span class="sxs-lookup"><span data-stu-id="615c9-244">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="615c9-245">У элементов **ExtensionPoint** этого типа может быть только один дочерний элемент **Group**.</span><span class="sxs-lookup"><span data-stu-id="615c9-245">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="615c9-246">Для атрибута **xsi:type** элементов **Control**, содержащихся в этой точке расширения, должно быть назначено значение `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="615c9-246">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="615c9-247">Пример</span><span class="sxs-lookup"><span data-stu-id="615c9-247">Example</span></span>
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

### <a name="events"></a><span data-ttu-id="615c9-248">События</span><span class="sxs-lookup"><span data-stu-id="615c9-248">Events</span></span>

<span data-ttu-id="615c9-249">Эта точка расширения добавляет обработчик для указанного события.</span><span class="sxs-lookup"><span data-stu-id="615c9-249">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="615c9-250">Этот тип элемента поддерживается классической версией Outlook в Интернете, доступен в [предварительной версии](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Windows и Mac, а также современной версии Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="615c9-250">This element type is supported by classic Outlook on the web, and in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Windows, Mac, and modern Outlook on the web.</span></span> <span data-ttu-id="615c9-251">Также требуется подписка на Office 365.</span><span class="sxs-lookup"><span data-stu-id="615c9-251">An Office 365 subscription is also required.</span></span>

| <span data-ttu-id="615c9-252">Элемент</span><span class="sxs-lookup"><span data-stu-id="615c9-252">Element</span></span> | <span data-ttu-id="615c9-253">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-253">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="615c9-254">Event</span><span class="sxs-lookup"><span data-stu-id="615c9-254">Event</span></span>](event.md) |  <span data-ttu-id="615c9-255">Задает событие и функцию его обработчика.</span><span class="sxs-lookup"><span data-stu-id="615c9-255">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="615c9-256">Пример события ItemSend</span><span class="sxs-lookup"><span data-stu-id="615c9-256">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="615c9-257">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="615c9-257">DetectedEntity</span></span>

<span data-ttu-id="615c9-258">Эта точка расширения добавляет активацию контекстной надстройки для указанного типа сущности.</span><span class="sxs-lookup"><span data-stu-id="615c9-258">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="615c9-259">В соответствующем элементе [VersionOverrides](versionoverrides.md) для атрибута `xsi:type` должно быть задано значение `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="615c9-259">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="615c9-260">Этот тип элемента доступен в [клиентах Outlook, поддерживающих наборы обязательных требований 1.6 и более поздних версий.](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)</span><span class="sxs-lookup"><span data-stu-id="615c9-260">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="615c9-261">Элемент</span><span class="sxs-lookup"><span data-stu-id="615c9-261">Element</span></span> |  <span data-ttu-id="615c9-262">Описание</span><span class="sxs-lookup"><span data-stu-id="615c9-262">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="615c9-263">Label</span><span class="sxs-lookup"><span data-stu-id="615c9-263">Label</span></span>](#label) |  <span data-ttu-id="615c9-264">Задает метку для надстройки в контекстном окне.</span><span class="sxs-lookup"><span data-stu-id="615c9-264">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="615c9-265">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="615c9-265">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="615c9-266">Задает URL-адрес контекстного окна.</span><span class="sxs-lookup"><span data-stu-id="615c9-266">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="615c9-267">Rule</span><span class="sxs-lookup"><span data-stu-id="615c9-267">Rule</span></span>](rule.md) |  <span data-ttu-id="615c9-268">Задает одно или несколько правил, определяющих, когда активируется надстройка.</span><span class="sxs-lookup"><span data-stu-id="615c9-268">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="615c9-269">Label</span><span class="sxs-lookup"><span data-stu-id="615c9-269">Label</span></span>

<span data-ttu-id="615c9-p116">Обязательно. Метка группы. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="615c9-p116">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="615c9-273">Требования к выделению</span><span class="sxs-lookup"><span data-stu-id="615c9-273">Highlight requirements</span></span>

<span data-ttu-id="615c9-p117">Единственный способ, которым пользователь может активировать контекстную надстройку, — взаимодействие с выделенной сущностью. Разработчики могут указывать, какие сущности выделяются, с помощью атрибута `Highlight` элемента `Rule` для типов правил `ItemHasKnownEntity` и `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="615c9-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="615c9-p118">Однако следует учитывать некоторые ограничения. Они гарантируют, что в соответствующих сообщениях и встречах всегда есть выделенная сущность, с помощью которой пользователь может активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="615c9-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="615c9-278">Сущности `EmailAddress` и `Url` не поддерживают выделение, поэтому их нельзя использовать для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="615c9-278">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="615c9-279">Если используется одно правило, то для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="615c9-279">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="615c9-280">Если используется правило `RuleCollection`, совмещенное с другими правилами с помощью оператора `Mode="AND"`, то как минимум в одном из правил для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="615c9-280">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="615c9-281">Если используется правило `RuleCollection`, в котором правила совмещаются с помощью оператора `Mode="OR"`, то в каждом из них для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="615c9-281">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="615c9-282">Пример события DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="615c9-282">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```
