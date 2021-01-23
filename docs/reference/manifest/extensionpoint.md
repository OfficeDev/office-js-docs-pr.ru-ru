---
title: Элемент ExtensionPoint в файле манифеста
description: Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.
ms.date: 01/22/2021
localization_priority: Normal
ms.openlocfilehash: 96bf3a6835b1a0ab6e5aa85a837515a3071e5610
ms.sourcegitcommit: 6c5716d92312887e3d944bf12d9985560109b3c0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2021
ms.locfileid: "49944307"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="4994a-103">Элемент ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="4994a-103">ExtensionPoint element</span></span>

 <span data-ttu-id="4994a-104">Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="4994a-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="4994a-105">Элемент **ExtensionPoint** является дочерним для элемента [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="4994a-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="4994a-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4994a-106">Attributes</span></span>

|  <span data-ttu-id="4994a-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="4994a-107">Attribute</span></span>  |  <span data-ttu-id="4994a-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4994a-108">Required</span></span>  |  <span data-ttu-id="4994a-109">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4994a-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="4994a-110">**xsi:type**</span></span>  |  <span data-ttu-id="4994a-111">Да</span><span class="sxs-lookup"><span data-stu-id="4994a-111">Yes</span></span>  | <span data-ttu-id="4994a-112">Тип определяемой точки расширения.</span><span class="sxs-lookup"><span data-stu-id="4994a-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="4994a-113">Точки расширения только для Excel</span><span class="sxs-lookup"><span data-stu-id="4994a-113">Extension points for Excel only</span></span>

- <span data-ttu-id="4994a-114">**CustomFunctions** — пользовательская функция, написанная на JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="4994a-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="4994a-115">[В этом примере кода XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) показано, как использовать элемент **ExtensionPoint** со значением атрибута **CustomFunctions** и какие дочерние элементы следует использовать.</span><span class="sxs-lookup"><span data-stu-id="4994a-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="4994a-116">Точки расширения для команд надстроек Word, Excel, PowerPoint и OneNote</span><span class="sxs-lookup"><span data-stu-id="4994a-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="4994a-117">**PrimaryCommandSurface** — лента в Office.</span><span class="sxs-lookup"><span data-stu-id="4994a-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="4994a-118">**ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="4994a-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="4994a-119">В приведенных ниже примерах показано, как применять элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.</span><span class="sxs-lookup"><span data-stu-id="4994a-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4994a-p102">Для элементов, которые содержат атрибут ID, обязательно предоставляйте уникальный идентификатор. Мы рекомендуем использовать название вашей компании и личный идентификатор. Пример формата приведен ниже. <CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="4994a-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="4994a-123">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-123">Child elements</span></span>
 
|<span data-ttu-id="4994a-124">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-124">Element</span></span>|<span data-ttu-id="4994a-125">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-125">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="4994a-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="4994a-126">**CustomTab**</span></span>|<span data-ttu-id="4994a-p103">Обязательный, если требуется добавить пользовательскую вкладку в ленту (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **CustomTab** и **OfficeTab** одновременно. Атрибут **id** является обязательным. </span><span class="sxs-lookup"><span data-stu-id="4994a-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="4994a-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="4994a-130">**OfficeTab**</span></span>|<span data-ttu-id="4994a-131">Обязательно, если вы хотите расширить вкладку ленты приложения Office по умолчанию (с помощью **PrimaryCommandSurface).**</span><span class="sxs-lookup"><span data-stu-id="4994a-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="4994a-132">Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно.</span><span class="sxs-lookup"><span data-stu-id="4994a-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="4994a-133">Подробные сведения см. [в officeTab.](officetab.md)</span><span class="sxs-lookup"><span data-stu-id="4994a-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="4994a-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="4994a-134">**OfficeMenu**</span></span>|<span data-ttu-id="4994a-p105">Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: </span><span class="sxs-lookup"><span data-stu-id="4994a-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="4994a-p106">- **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши. </span><span class="sxs-lookup"><span data-stu-id="4994a-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="4994a-p107">- **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь нажимает ячейку электронной таблицы правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="4994a-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="4994a-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="4994a-141">**Group**</span></span>|<span data-ttu-id="4994a-p108">Группа точек расширения интерфейса пользователя на вкладке. В группе может быть до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов. </span><span class="sxs-lookup"><span data-stu-id="4994a-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="4994a-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="4994a-145">**Label**</span></span>|<span data-ttu-id="4994a-146">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="4994a-146">Required.</span></span> <span data-ttu-id="4994a-147">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="4994a-147">The label of the group.</span></span> <span data-ttu-id="4994a-148">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String.**</span><span class="sxs-lookup"><span data-stu-id="4994a-148">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="4994a-149">**String** — это дочерний элемент **ShortStrings**, который в свою очередь является дочерним для элемента **Resources**.</span><span class="sxs-lookup"><span data-stu-id="4994a-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="4994a-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="4994a-150">**Icon**</span></span>|<span data-ttu-id="4994a-151">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="4994a-151">Required.</span></span> <span data-ttu-id="4994a-152">Определяет значок группы для использования на устройствах с малым форм-фактором или в случаях, когда отображается слишком много кнопок.</span><span class="sxs-lookup"><span data-stu-id="4994a-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="4994a-153">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **Image.**</span><span class="sxs-lookup"><span data-stu-id="4994a-153">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="4994a-154">**Image** — это дочерний элемент **Images**, который в свою очередь является дочерним для элемента **Resources**.</span><span class="sxs-lookup"><span data-stu-id="4994a-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="4994a-155">Атрибут **size** определяет размер изображения в пикселях.</span><span class="sxs-lookup"><span data-stu-id="4994a-155">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="4994a-156">Обязательными являются три размера изображения: 16, 32 и 80.</span><span class="sxs-lookup"><span data-stu-id="4994a-156">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="4994a-157">Кроме того, поддерживаются пять необязательных размеров: 20, 24, 40, 48 и 64.</span><span class="sxs-lookup"><span data-stu-id="4994a-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="4994a-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="4994a-158">**Tooltip**</span></span>|<span data-ttu-id="4994a-159">Необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="4994a-159">Optional.</span></span> <span data-ttu-id="4994a-160">Всплывающая подсказка группы.</span><span class="sxs-lookup"><span data-stu-id="4994a-160">The tooltip of the group.</span></span> <span data-ttu-id="4994a-161">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String.**</span><span class="sxs-lookup"><span data-stu-id="4994a-161">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="4994a-162">**String** — это дочерний элемент **LongStrings**, который в свою очередь является дочерним для элемента **Resources**.</span><span class="sxs-lookup"><span data-stu-id="4994a-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="4994a-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="4994a-163">**Control**</span></span>|<span data-ttu-id="4994a-164">В каждой группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="4994a-164">Each group requires at least one control.</span></span> <span data-ttu-id="4994a-165">Элемент **управления** может быть кнопкой **или** **меню.**</span><span class="sxs-lookup"><span data-stu-id="4994a-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="4994a-166">Используйте **меню** для указания выпадаемого списка элементов управления кнопками.</span><span class="sxs-lookup"><span data-stu-id="4994a-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="4994a-167">В настоящее время поддерживаются только кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="4994a-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="4994a-168">Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](control.md#button-control) и [Элементы управления меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="4994a-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="4994a-169">**Примечание.**  Чтобы упростить устранение неполадок, рекомендуется добавлять по одному **элементу Control** и связанным с ним элементам **Resources.**</span><span class="sxs-lookup"><span data-stu-id="4994a-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="4994a-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="4994a-170">**Script**</span></span>|<span data-ttu-id="4994a-171">Ссылка на файл JavaScript с пользовательским определением функции и кодом регистрации.</span><span class="sxs-lookup"><span data-stu-id="4994a-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="4994a-172">Этот элемент не используется в предварительной версии для разработчиков.</span><span class="sxs-lookup"><span data-stu-id="4994a-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="4994a-173">Загрузку всех файлов JavaScript выполняет страница HTML.</span><span class="sxs-lookup"><span data-stu-id="4994a-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="4994a-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="4994a-174">**Page**</span></span>|<span data-ttu-id="4994a-175">Ссылка на HTML-страницу для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="4994a-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="4994a-176">Точки расширения для Outlook</span><span class="sxs-lookup"><span data-stu-id="4994a-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="4994a-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="4994a-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="4994a-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="4994a-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="4994a-181">[Module](#module) (можно использовать только в [DesktopFormFactor](desktopformfactor.md))</span><span class="sxs-lookup"><span data-stu-id="4994a-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="4994a-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="4994a-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface)
- [<span data-ttu-id="4994a-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="4994a-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="4994a-185">Events</span><span class="sxs-lookup"><span data-stu-id="4994a-185">Events</span></span>](#events)
- [<span data-ttu-id="4994a-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="4994a-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="4994a-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="4994a-p114">Эта точка расширения помещает кнопки на панель команд для представления чтения почты. В классической версии Outlook эта панель отображается на ленте.</span><span class="sxs-lookup"><span data-stu-id="4994a-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4994a-190">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-190">Child elements</span></span>

|  <span data-ttu-id="4994a-191">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-191">Element</span></span> |  <span data-ttu-id="4994a-192">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4994a-194">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4994a-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4994a-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4994a-196">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="4994a-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="4994a-197">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="4994a-198">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="4994a-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="4994a-200">Эта точка расширения добавляет кнопки на ленту для надстроек, использующих форму создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="4994a-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="4994a-201">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-201">Child elements</span></span>

|  <span data-ttu-id="4994a-202">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-202">Element</span></span> |  <span data-ttu-id="4994a-203">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4994a-205">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4994a-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4994a-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4994a-207">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="4994a-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="4994a-208">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="4994a-209">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="4994a-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="4994a-211">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="4994a-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="4994a-212">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-212">Child elements</span></span>

|  <span data-ttu-id="4994a-213">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-213">Element</span></span> |  <span data-ttu-id="4994a-214">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4994a-216">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4994a-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4994a-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4994a-218">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="4994a-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="4994a-219">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="4994a-220">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="4994a-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="4994a-222">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для участника собрания.</span><span class="sxs-lookup"><span data-stu-id="4994a-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="4994a-223">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-223">Child elements</span></span>

|  <span data-ttu-id="4994a-224">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-224">Element</span></span> |  <span data-ttu-id="4994a-225">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4994a-227">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4994a-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4994a-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4994a-229">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="4994a-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="4994a-230">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="4994a-231">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="4994a-232">Module</span><span class="sxs-lookup"><span data-stu-id="4994a-232">Module</span></span>

<span data-ttu-id="4994a-233">Эта точка расширения добавляет кнопки на ленту для расширения модуля.</span><span class="sxs-lookup"><span data-stu-id="4994a-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4994a-234">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-234">Child elements</span></span>

|  <span data-ttu-id="4994a-235">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-235">Element</span></span> |  <span data-ttu-id="4994a-236">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-236">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-237">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4994a-237">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4994a-238">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4994a-238">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4994a-239">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4994a-239">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4994a-240">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="4994a-240">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="4994a-241">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-241">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="4994a-242">Эта точка расширения помещает кнопки на панель команд для чтения почты в форм-факторе мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="4994a-242">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4994a-243">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-243">Child elements</span></span>

|  <span data-ttu-id="4994a-244">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-244">Element</span></span> |  <span data-ttu-id="4994a-245">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-245">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-246">Group</span><span class="sxs-lookup"><span data-stu-id="4994a-246">Group</span></span>](group.md) |  <span data-ttu-id="4994a-247">Добавляет группу кнопок на панель команд.</span><span class="sxs-lookup"><span data-stu-id="4994a-247">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="4994a-248">У элементов **ExtensionPoint** этого типа может быть только один дочерний элемент **Group**.</span><span class="sxs-lookup"><span data-stu-id="4994a-248">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="4994a-249">Для атрибута **xsi:type** элементов **Control**, содержащихся в этой точке расширения, должно быть назначено значение `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="4994a-249">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="4994a-250">Пример</span><span class="sxs-lookup"><span data-stu-id="4994a-250">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface"></a><span data-ttu-id="4994a-251">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4994a-251">MobileOnlineMeetingCommandSurface</span></span>

<span data-ttu-id="4994a-252">Эта точка расширения помещает соответствующий режиму в командную поверхность для встречи в форм-факторе мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="4994a-252">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="4994a-253">Организатор собрания может создать собрание по сети.</span><span class="sxs-lookup"><span data-stu-id="4994a-253">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="4994a-254">После этого участник может присоединиться к собранию по сети.</span><span class="sxs-lookup"><span data-stu-id="4994a-254">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="4994a-255">Дополнительные информацию об этом сценарии см. в статье "Создание мобильной надстройки [Outlook"](../../outlook/online-meeting.md) для поставщика собраний по сети.</span><span class="sxs-lookup"><span data-stu-id="4994a-255">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

> [!NOTE]
> <span data-ttu-id="4994a-256">Эта точка расширения поддерживается только на Android с подпиской на Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="4994a-256">This extension point is only supported on Android with a Microsoft 365 subscription.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4994a-257">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-257">Child elements</span></span>

|  <span data-ttu-id="4994a-258">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-258">Element</span></span> |  <span data-ttu-id="4994a-259">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-259">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-260">Control</span><span class="sxs-lookup"><span data-stu-id="4994a-260">Control</span></span>](control.md) |  <span data-ttu-id="4994a-261">Добавляет кнопку на поверхность команды.</span><span class="sxs-lookup"><span data-stu-id="4994a-261">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="4994a-262">`ExtensionPoint` элементы этого типа могут иметь только один элемент: `Control` элемент.</span><span class="sxs-lookup"><span data-stu-id="4994a-262">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="4994a-263">Элемент, `Control` содержащийся в этой точке расширения, должен иметь `xsi:type` атрибут , установленный в `MobileButton` .</span><span class="sxs-lookup"><span data-stu-id="4994a-263">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="4994a-264">Изображения должны быть серыми с использованием `Icon` hex-кода или `#919191` их эквивалента в [других форматах цвета.](https://convertingcolors.com/hex-color-919191.html)</span><span class="sxs-lookup"><span data-stu-id="4994a-264">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="4994a-265">Пример</span><span class="sxs-lookup"><span data-stu-id="4994a-265">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="4994a-266">LaunchEvent (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="4994a-266">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="4994a-267">Эта точка расширения поддерживается только в предварительной [версии](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) Outlook в Интернете с подпиской на Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="4994a-267">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="4994a-268">Эта точка расширения позволяет надстройка активироваться на основе поддерживаемых событий в форм-факторе рабочего стола.</span><span class="sxs-lookup"><span data-stu-id="4994a-268">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="4994a-269">В настоящее время поддерживаются только события `OnNewMessageCompose` и `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="4994a-269">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="4994a-270">Дополнительные информацию об этом сценарии см. в статье "Настройка надстройки Outlook для [активации на](../../outlook/autolaunch.md) основе событий".</span><span class="sxs-lookup"><span data-stu-id="4994a-270">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4994a-271">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4994a-271">Child elements</span></span>

|  <span data-ttu-id="4994a-272">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-272">Element</span></span> |  <span data-ttu-id="4994a-273">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-273">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="4994a-274">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="4994a-274">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="4994a-275">Список [LaunchEvent для](launchevent.md) активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="4994a-275">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="4994a-276">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="4994a-276">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="4994a-277">Расположение исходных файлов JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4994a-277">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="4994a-278">Пример</span><span class="sxs-lookup"><span data-stu-id="4994a-278">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="4994a-279">События</span><span class="sxs-lookup"><span data-stu-id="4994a-279">Events</span></span>

<span data-ttu-id="4994a-280">Эта точка расширения добавляет обработчик для указанного события.</span><span class="sxs-lookup"><span data-stu-id="4994a-280">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="4994a-281">Дополнительные сведения об использовании этой точки расширения см. в функции отправки для [надстройки Outlook.](../../outlook/outlook-on-send-addins.md)</span><span class="sxs-lookup"><span data-stu-id="4994a-281">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

| <span data-ttu-id="4994a-282">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-282">Element</span></span> | <span data-ttu-id="4994a-283">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-283">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-284">Event</span><span class="sxs-lookup"><span data-stu-id="4994a-284">Event</span></span>](event.md) |  <span data-ttu-id="4994a-285">Задает событие и функцию его обработчика.</span><span class="sxs-lookup"><span data-stu-id="4994a-285">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="4994a-286">Пример события ItemSend</span><span class="sxs-lookup"><span data-stu-id="4994a-286">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="4994a-287">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="4994a-287">DetectedEntity</span></span>

<span data-ttu-id="4994a-288">Эта точка расширения добавляет активацию контекстной надстройки для указанного типа сущности.</span><span class="sxs-lookup"><span data-stu-id="4994a-288">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="4994a-289">В соответствующем элементе [VersionOverrides](versionoverrides.md) для атрибута `xsi:type` должно быть задано значение `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="4994a-289">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="4994a-290">Этот тип элемента доступен в [клиентах Outlook, поддерживающих наборы обязательных требований 1.6 и более поздних версий.](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)</span><span class="sxs-lookup"><span data-stu-id="4994a-290">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="4994a-291">Элемент</span><span class="sxs-lookup"><span data-stu-id="4994a-291">Element</span></span> |  <span data-ttu-id="4994a-292">Описание</span><span class="sxs-lookup"><span data-stu-id="4994a-292">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4994a-293">Label</span><span class="sxs-lookup"><span data-stu-id="4994a-293">Label</span></span>](#label) |  <span data-ttu-id="4994a-294">Задает метку для надстройки в контекстном окне.</span><span class="sxs-lookup"><span data-stu-id="4994a-294">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="4994a-295">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="4994a-295">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="4994a-296">Задает URL-адрес контекстного окна.</span><span class="sxs-lookup"><span data-stu-id="4994a-296">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="4994a-297">Rule</span><span class="sxs-lookup"><span data-stu-id="4994a-297">Rule</span></span>](rule.md) |  <span data-ttu-id="4994a-298">Задает одно или несколько правил, определяющих, когда активируется надстройка.</span><span class="sxs-lookup"><span data-stu-id="4994a-298">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="4994a-299">Label</span><span class="sxs-lookup"><span data-stu-id="4994a-299">Label</span></span>

<span data-ttu-id="4994a-300">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="4994a-300">Required.</span></span> <span data-ttu-id="4994a-301">Метка группы.</span><span class="sxs-lookup"><span data-stu-id="4994a-301">The label of the group.</span></span> <span data-ttu-id="4994a-302">Атрибут **resid** может быть не более 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="4994a-302">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="4994a-303">Требования к выделению</span><span class="sxs-lookup"><span data-stu-id="4994a-303">Highlight requirements</span></span>

<span data-ttu-id="4994a-p119">Единственный способ, которым пользователь может активировать контекстную надстройку, — взаимодействие с выделенной сущностью. Разработчики могут указывать, какие сущности выделяются, с помощью атрибута `Highlight` элемента `Rule` для типов правил `ItemHasKnownEntity` и `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="4994a-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="4994a-p120">Однако следует учитывать некоторые ограничения. Они гарантируют, что в соответствующих сообщениях и встречах всегда есть выделенная сущность, с помощью которой пользователь может активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="4994a-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="4994a-308">Сущности `EmailAddress` и `Url` не поддерживают выделение, поэтому их нельзя использовать для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="4994a-308">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="4994a-309">Если используется одно правило, то для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="4994a-309">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="4994a-310">Если используется правило `RuleCollection`, совмещенное с другими правилами с помощью оператора `Mode="AND"`, то как минимум в одном из правил для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="4994a-310">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="4994a-311">Если используется правило `RuleCollection`, в котором правила совмещаются с помощью оператора `Mode="OR"`, то в каждом из них для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="4994a-311">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="4994a-312">Пример события DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="4994a-312">DetectedEntity event example</span></span>

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
