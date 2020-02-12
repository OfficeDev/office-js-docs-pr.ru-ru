---
title: Элемент ExtensionPoint в файле манифеста
description: ''
ms.date: 09/05/2019
localization_priority: Normal
ms.openlocfilehash: 2ad9e0ccb0393e562ca0bb6951ec9bc4eb9c3eab
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950707"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="9c10a-102">Элемент ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="9c10a-102">ExtensionPoint element</span></span>

 <span data-ttu-id="9c10a-103">Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="9c10a-103">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="9c10a-104">Элемент **ExtensionPoint** является дочерним для элемента [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="9c10a-104">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="9c10a-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9c10a-105">Attributes</span></span>

|  <span data-ttu-id="9c10a-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="9c10a-106">Attribute</span></span>  |  <span data-ttu-id="9c10a-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9c10a-107">Required</span></span>  |  <span data-ttu-id="9c10a-108">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9c10a-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="9c10a-109">**xsi:type**</span></span>  |  <span data-ttu-id="9c10a-110">Да</span><span class="sxs-lookup"><span data-stu-id="9c10a-110">Yes</span></span>  | <span data-ttu-id="9c10a-111">Тип определяемой точки расширения.</span><span class="sxs-lookup"><span data-stu-id="9c10a-111">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="9c10a-112">Точки расширения только для Excel</span><span class="sxs-lookup"><span data-stu-id="9c10a-112">Extension points for Excel only</span></span>

- <span data-ttu-id="9c10a-113">**CustomFunctions** — пользовательская функция, написанная на JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="9c10a-113">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="9c10a-114">[В этом примере кода XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) показано, как использовать элемент **ExtensionPoint** со значением атрибута **CustomFunctions** и какие дочерние элементы следует использовать.</span><span class="sxs-lookup"><span data-stu-id="9c10a-114">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="9c10a-115">Точки расширения для команд надстроек Word, Excel, PowerPoint и OneNote</span><span class="sxs-lookup"><span data-stu-id="9c10a-115">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="9c10a-116">**PrimaryCommandSurface** — лента в Office.</span><span class="sxs-lookup"><span data-stu-id="9c10a-116">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="9c10a-117">**ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="9c10a-117">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="9c10a-118">В следующих примерах показано, как использовать элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.</span><span class="sxs-lookup"><span data-stu-id="9c10a-118">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="9c10a-p102">Для элементов, которые содержат атрибут ID, обязательно предоставляйте уникальный идентификатор. Мы рекомендуем использовать название вашей компании и личный идентификатор. Пример формата приведен ниже. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="9c10a-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="9c10a-122">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9c10a-122">Child elements</span></span>
 
|<span data-ttu-id="9c10a-123">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="9c10a-123">**Element**</span></span>|<span data-ttu-id="9c10a-124">**Описание**</span><span class="sxs-lookup"><span data-stu-id="9c10a-124">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9c10a-125">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="9c10a-125">**CustomTab**</span></span>|<span data-ttu-id="9c10a-p103">Обязательный, если требуется добавить на ленту настраиваемую вкладку (с помощью элемента **PrimaryCommandSurface**). Если используется элемент **CustomTab**, использовать элемент **OfficeTab** невозможно. Атрибут **id** является обязательным.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="9c10a-129">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="9c10a-129">**OfficeTab**</span></span>|<span data-ttu-id="9c10a-p104">Обязательный, если требуется расширить стандартную вкладку ленты Office (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно. Дополнительные сведения см. в статье [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="9c10a-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="9c10a-133">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="9c10a-133">**OfficeMenu**</span></span>|<span data-ttu-id="9c10a-p105">Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: </span><span class="sxs-lookup"><span data-stu-id="9c10a-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="9c10a-p106">- **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши. </span><span class="sxs-lookup"><span data-stu-id="9c10a-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="9c10a-p107">- **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь нажимает ячейку электронной таблицы правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="9c10a-140">**Group**</span><span class="sxs-lookup"><span data-stu-id="9c10a-140">**Group**</span></span>|<span data-ttu-id="9c10a-p108">Группа точек расширения пользовательского интерфейса на вкладке. Группа может включать до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="9c10a-144">**Label**</span><span class="sxs-lookup"><span data-stu-id="9c10a-144">**Label**</span></span>|<span data-ttu-id="9c10a-p109">Обязательный. Метка группы. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **String**. Элемент **String** — это дочерний элемент элемента **ShortStrings**, который является дочерним для элемента **Resources**.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="9c10a-149">**Icon**</span><span class="sxs-lookup"><span data-stu-id="9c10a-149">**Icon**</span></span>|<span data-ttu-id="9c10a-p110">Обязательный. Задает значок группы, который будет использоваться на устройствах с малым форм-фактором либо при отображении слишком большого количества кнопок. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **Image**. Элемент **Image** — это дочерний элемент элемента **Images**, который является дочерним для элемента **Resources**. Атрибут **size** указывает размер изображения в пикселях. Необходимо три размера изображения: 16, 32 и 80. Кроме того, поддерживается пять необязательных размеров: 20, 24, 40, 48 и 64.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="9c10a-157">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="9c10a-157">**Tooltip**</span></span>|<span data-ttu-id="9c10a-p111">Необязательный. Подсказка группы. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **String**. Элемент **String** — это дочерний элемент элемента **LongStrings**, который является дочерним для элемента **Resources**.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="9c10a-162">**Control**</span><span class="sxs-lookup"><span data-stu-id="9c10a-162">**Control**</span></span>|<span data-ttu-id="9c10a-163">В каждой группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="9c10a-163">Each group requires at least one control.</span></span> <span data-ttu-id="9c10a-164">Элемент **Control** может относиться к типу **Button** или **Menu**.</span><span class="sxs-lookup"><span data-stu-id="9c10a-164">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="9c10a-165">С помощью элемента **Menu** можно указать раскрывающийся список элементов управления "Кнопка".</span><span class="sxs-lookup"><span data-stu-id="9c10a-165">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="9c10a-166">В настоящее время поддерживаются только кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="9c10a-166">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="9c10a-167">Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](control.md#button-control) и [Элементы управления меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="9c10a-167">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="9c10a-168">**Примечание.** Чтобы упростить устранение неполадок, рекомендуем добавлять элемент **Control** и соответствующий дочерний элемент **Resources** по одному.</span><span class="sxs-lookup"><span data-stu-id="9c10a-168">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="9c10a-169">**Script**</span><span class="sxs-lookup"><span data-stu-id="9c10a-169">**Script**</span></span>|<span data-ttu-id="9c10a-170">Ссылка на файл JavaScript с пользовательским определением функции и кодом регистрации.</span><span class="sxs-lookup"><span data-stu-id="9c10a-170">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="9c10a-171">Этот элемент не используется в предварительной версии для разработчиков.</span><span class="sxs-lookup"><span data-stu-id="9c10a-171">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="9c10a-172">Загрузку всех файлов JavaScript выполняет страница HTML.</span><span class="sxs-lookup"><span data-stu-id="9c10a-172">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="9c10a-173">**Page**</span><span class="sxs-lookup"><span data-stu-id="9c10a-173">**Page**</span></span>|<span data-ttu-id="9c10a-174">Ссылка на HTML-страницу для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="9c10a-174">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="9c10a-175">Точки расширения для Outlook</span><span class="sxs-lookup"><span data-stu-id="9c10a-175">Extension points for Outlook</span></span>

- [<span data-ttu-id="9c10a-176">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-176">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="9c10a-177">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-177">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="9c10a-178">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-178">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="9c10a-179">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-179">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="9c10a-180">[Module](#module) (можно использовать только в [DesktopFormFactor](desktopformfactor.md))</span><span class="sxs-lookup"><span data-stu-id="9c10a-180">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="9c10a-181">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-181">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="9c10a-182">Events</span><span class="sxs-lookup"><span data-stu-id="9c10a-182">Events</span></span>](#events)
- [<span data-ttu-id="9c10a-183">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="9c10a-183">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="9c10a-184">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-184">MessageReadCommandSurface</span></span>
<span data-ttu-id="9c10a-p114">Эта точка расширения помещает кнопки на панель команд для представления чтения почты. В классической версии Outlook эта панель отображается на ленте.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9c10a-187">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9c10a-187">Child elements</span></span>

|  <span data-ttu-id="9c10a-188">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c10a-188">Element</span></span> |  <span data-ttu-id="9c10a-189">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-189">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c10a-190">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-190">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c10a-191">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9c10a-191">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c10a-192">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-192">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c10a-193">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="9c10a-193">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9c10a-194">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-194">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9c10a-195">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-195">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="9c10a-196">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-196">MessageComposeCommandSurface</span></span>
<span data-ttu-id="9c10a-197">Эта точка расширения добавляет кнопки на ленту для надстроек, использующих форму создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="9c10a-197">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9c10a-198">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9c10a-198">Child elements</span></span>

|  <span data-ttu-id="9c10a-199">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c10a-199">Element</span></span> |  <span data-ttu-id="9c10a-200">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-200">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c10a-201">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-201">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c10a-202">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9c10a-202">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c10a-203">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-203">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c10a-204">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="9c10a-204">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9c10a-205">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-205">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9c10a-206">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-206">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="9c10a-207">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-207">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="9c10a-208">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="9c10a-208">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9c10a-209">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9c10a-209">Child elements</span></span>

|  <span data-ttu-id="9c10a-210">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c10a-210">Element</span></span> |  <span data-ttu-id="9c10a-211">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-211">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c10a-212">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-212">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c10a-213">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9c10a-213">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c10a-214">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-214">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c10a-215">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="9c10a-215">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9c10a-216">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-216">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9c10a-217">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-217">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="9c10a-218">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-218">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="9c10a-219">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для участника собрания.</span><span class="sxs-lookup"><span data-stu-id="9c10a-219">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9c10a-220">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9c10a-220">Child elements</span></span>

|  <span data-ttu-id="9c10a-221">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c10a-221">Element</span></span> |  <span data-ttu-id="9c10a-222">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-222">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c10a-223">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-223">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c10a-224">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9c10a-224">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c10a-225">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-225">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c10a-226">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="9c10a-226">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9c10a-227">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-227">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9c10a-228">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-228">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="9c10a-229">Module</span><span class="sxs-lookup"><span data-stu-id="9c10a-229">Module</span></span>

<span data-ttu-id="9c10a-230">Эта точка расширения добавляет кнопки на ленту для расширения модуля.</span><span class="sxs-lookup"><span data-stu-id="9c10a-230">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9c10a-231">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9c10a-231">Child elements</span></span>

|  <span data-ttu-id="9c10a-232">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c10a-232">Element</span></span> |  <span data-ttu-id="9c10a-233">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-233">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c10a-234">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-234">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9c10a-235">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9c10a-235">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9c10a-236">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9c10a-236">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9c10a-237">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="9c10a-237">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="9c10a-238">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9c10a-238">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="9c10a-239">Эта точка расширения помещает кнопки на панель команд для чтения почты в форм-факторе мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="9c10a-239">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9c10a-240">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9c10a-240">Child elements</span></span>

|  <span data-ttu-id="9c10a-241">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c10a-241">Element</span></span> |  <span data-ttu-id="9c10a-242">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-242">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c10a-243">Group</span><span class="sxs-lookup"><span data-stu-id="9c10a-243">Group</span></span>](group.md) |  <span data-ttu-id="9c10a-244">Добавляет группу кнопок на панель команд.</span><span class="sxs-lookup"><span data-stu-id="9c10a-244">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="9c10a-245">У элементов **ExtensionPoint** этого типа может быть только один дочерний элемент **Group**.</span><span class="sxs-lookup"><span data-stu-id="9c10a-245">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="9c10a-246">Для атрибута **xsi:type** элементов **Control**, содержащихся в этой точке расширения, должно быть назначено значение `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="9c10a-246">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="9c10a-247">Пример</span><span class="sxs-lookup"><span data-stu-id="9c10a-247">Example</span></span>
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

### <a name="events"></a><span data-ttu-id="9c10a-248">События</span><span class="sxs-lookup"><span data-stu-id="9c10a-248">Events</span></span>

<span data-ttu-id="9c10a-249">Эта точка расширения добавляет обработчик для указанного события.</span><span class="sxs-lookup"><span data-stu-id="9c10a-249">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="9c10a-250">Этот тип элемента поддерживается классической версией Outlook в Интернете, доступен в [предварительной версии](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Windows и Mac, а также современной версии Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="9c10a-250">This element type is supported by classic Outlook on the web, and in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Windows, Mac, and modern Outlook on the web.</span></span> <span data-ttu-id="9c10a-251">Также требуется подписка на Office 365.</span><span class="sxs-lookup"><span data-stu-id="9c10a-251">An Office 365 subscription is also required.</span></span>

| <span data-ttu-id="9c10a-252">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c10a-252">Element</span></span> | <span data-ttu-id="9c10a-253">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-253">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c10a-254">Event</span><span class="sxs-lookup"><span data-stu-id="9c10a-254">Event</span></span>](event.md) |  <span data-ttu-id="9c10a-255">Задает событие и функцию его обработчика.</span><span class="sxs-lookup"><span data-stu-id="9c10a-255">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="9c10a-256">Пример события ItemSend</span><span class="sxs-lookup"><span data-stu-id="9c10a-256">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="9c10a-257">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="9c10a-257">DetectedEntity</span></span>

<span data-ttu-id="9c10a-258">Эта точка расширения добавляет активацию контекстной надстройки для указанного типа сущности.</span><span class="sxs-lookup"><span data-stu-id="9c10a-258">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="9c10a-259">В соответствующем элементе [VersionOverrides](versionoverrides.md) для атрибута `xsi:type` должно быть задано значение `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="9c10a-259">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="9c10a-260">Этот тип элемента доступен в [клиентах Outlook, поддерживающих наборы обязательных требований 1.6 и более поздних версий.](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)</span><span class="sxs-lookup"><span data-stu-id="9c10a-260">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="9c10a-261">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c10a-261">Element</span></span> |  <span data-ttu-id="9c10a-262">Описание</span><span class="sxs-lookup"><span data-stu-id="9c10a-262">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9c10a-263">Label</span><span class="sxs-lookup"><span data-stu-id="9c10a-263">Label</span></span>](#label) |  <span data-ttu-id="9c10a-264">Задает метку для надстройки в контекстном окне.</span><span class="sxs-lookup"><span data-stu-id="9c10a-264">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="9c10a-265">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9c10a-265">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="9c10a-266">Задает URL-адрес контекстного окна.</span><span class="sxs-lookup"><span data-stu-id="9c10a-266">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="9c10a-267">Rule</span><span class="sxs-lookup"><span data-stu-id="9c10a-267">Rule</span></span>](rule.md) |  <span data-ttu-id="9c10a-268">Задает одно или несколько правил, определяющих, когда активируется надстройка.</span><span class="sxs-lookup"><span data-stu-id="9c10a-268">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="9c10a-269">Label</span><span class="sxs-lookup"><span data-stu-id="9c10a-269">Label</span></span>

<span data-ttu-id="9c10a-p116">Обязательный элемент. Метка группы. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="9c10a-p116">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="9c10a-273">Требования к выделению</span><span class="sxs-lookup"><span data-stu-id="9c10a-273">Highlight requirements</span></span>

<span data-ttu-id="9c10a-p117">Единственный способ, которым пользователь может активировать контекстную надстройку, — взаимодействие с выделенной сущностью. Разработчики могут указывать, какие сущности выделяются, с помощью атрибута `Highlight` элемента `Rule` для типов правил `ItemHasKnownEntity` и `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="9c10a-p118">Однако следует учитывать некоторые ограничения. Они гарантируют, что в соответствующих сообщениях и встречах всегда есть выделенная сущность, с помощью которой пользователь может активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="9c10a-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="9c10a-278">Сущности `EmailAddress` и `Url` не поддерживают выделение, поэтому их нельзя использовать для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c10a-278">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="9c10a-279">Если используется одно правило, то для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="9c10a-279">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="9c10a-280">Если используется правило `RuleCollection`, совмещенное с другими правилами с помощью оператора `Mode="AND"`, то как минимум в одном из правил для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="9c10a-280">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="9c10a-281">Если используется правило `RuleCollection`, в котором правила совмещаются с помощью оператора `Mode="OR"`, то в каждом из них для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="9c10a-281">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="9c10a-282">Пример события DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="9c10a-282">DetectedEntity event example</span></span>

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
