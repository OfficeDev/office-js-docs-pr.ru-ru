---
title: Интеграция встроенных кнопок Office в настраиваемые группы управления и вкладки
description: Узнайте, как включить встроенные кнопки Office в настраиваемые группы команд и вкладки на ленте Office.
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 8d4e8f39313551d001669b948b146250114f3e06
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505257"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a><span data-ttu-id="f38fd-103">Интеграция встроенных кнопок Office в настраиваемые группы управления и вкладки</span><span class="sxs-lookup"><span data-stu-id="f38fd-103">Integrate built-in Office buttons into custom control groups and tabs</span></span>

<span data-ttu-id="f38fd-104">Встроенные кнопки Office можно вставить в настраиваемые группы управления на ленте Office с помощью разметки в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="f38fd-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="f38fd-105">(Вы не можете вставить настраиваемые команды надстройки в встроенную группу Office.) Вы также можете вставить целые встроенные группы управления Office в пользовательские вкладки ленты.</span><span class="sxs-lookup"><span data-stu-id="f38fd-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="f38fd-106">В этой статье предполагается, что вы знакомы со статьей Основные понятия для команд [надстройки](add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="f38fd-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="f38fd-107">Пожалуйста, просмотрите его, если вы еще не сделали этого в последнее время.</span><span class="sxs-lookup"><span data-stu-id="f38fd-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="f38fd-108">Функция надстройки и разметка, описанные в этой статье, доступна только *в PowerPoint в Интернете.*</span><span class="sxs-lookup"><span data-stu-id="f38fd-108">The add-in feature and markup described in this article is *only available in PowerPoint on the web*.</span></span>
> - <span data-ttu-id="f38fd-109">Разметка, описанная в этой статье, работает только на платформах, поддерживаюх набор **требований AddinCommands 1.3.**</span><span class="sxs-lookup"><span data-stu-id="f38fd-109">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="f38fd-110">См. в более позднем [разделе Поведение на неподтверченных платформах.](#behavior-on-unsupported-platforms)</span><span class="sxs-lookup"><span data-stu-id="f38fd-110">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="f38fd-111">Вставьте встроенную группу управления в настраиваемую вкладку</span><span class="sxs-lookup"><span data-stu-id="f38fd-111">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="f38fd-112">Чтобы вставить встроенную группу управления Office в вкладку, добавьте элемент [OfficeGroup](../reference/manifest/customtab.md#officegroup) в родительский `<CustomTab>` элемент.</span><span class="sxs-lookup"><span data-stu-id="f38fd-112">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="f38fd-113">Атрибут `id` элемента `<OfficeGroup>` задается iD встроенной группы.</span><span class="sxs-lookup"><span data-stu-id="f38fd-113">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="f38fd-114">См. [в рублях Find the IDs of controls and control groups.](#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="f38fd-114">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="f38fd-115">В следующем примере разметки группа управления Пунктом Office добавляется в настраиваемую вкладку и позиционет ее, чтобы она появляться сразу после настраиваемой группы.</span><span class="sxs-lookup"><span data-stu-id="f38fd-115">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="f38fd-116">Вставьте встроенный контроль в настраиваемую группу</span><span class="sxs-lookup"><span data-stu-id="f38fd-116">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="f38fd-117">Чтобы вставить встроенный элемент управления Office в настраиваемую группу, добавьте элемент [OfficeControl](../reference/manifest/group.md#officecontrol) в родительский `<Group>` элемент.</span><span class="sxs-lookup"><span data-stu-id="f38fd-117">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="f38fd-118">Атрибут `id` элемента `<OfficeControl>` задается iD встроенного элемента управления.</span><span class="sxs-lookup"><span data-stu-id="f38fd-118">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="f38fd-119">См. [в рублях Find the IDs of controls and control groups.](#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="f38fd-119">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="f38fd-120">В следующем примере разметки управление Office Superscript добавляется в настраиваемую группу и должно отображаться сразу после настраиваемой кнопки.</span><span class="sxs-lookup"><span data-stu-id="f38fd-120">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.grp1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

> [!NOTE]
> <span data-ttu-id="f38fd-121">Пользователи могут настроить ленту в приложении Office.</span><span class="sxs-lookup"><span data-stu-id="f38fd-121">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="f38fd-122">Любые настройки пользователя переопределяют параметры манифеста.</span><span class="sxs-lookup"><span data-stu-id="f38fd-122">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="f38fd-123">Например, пользователь может удалить кнопку из любой группы и удалить любую группу со вкладки.</span><span class="sxs-lookup"><span data-stu-id="f38fd-123">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="f38fd-124">Поиск ID-элементов групп управления и управления</span><span class="sxs-lookup"><span data-stu-id="f38fd-124">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="f38fd-125">ID для поддерживаемых групп управления и управления находятся в файлах в ID-файлах управления репо [Office.](https://github.com/OfficeDev/office-control-ids)</span><span class="sxs-lookup"><span data-stu-id="f38fd-125">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="f38fd-126">Следуйте инструкциям в файле ReadMe этого репо.</span><span class="sxs-lookup"><span data-stu-id="f38fd-126">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="f38fd-127">Поведение на неподтверченных платформах</span><span class="sxs-lookup"><span data-stu-id="f38fd-127">Behavior on unsupported platforms</span></span>

<span data-ttu-id="f38fd-128">Если надстройка установлена на платформе, которая не поддерживает набор требований [AddinCommands 1.3,](../reference/requirement-sets/add-in-commands-requirement-sets.md)то разметка, описанная в этой статье, игнорируется, а встроенные элементы управления и группы Office не будут отображаться в настраиваемой группе или вкладке.</span><span class="sxs-lookup"><span data-stu-id="f38fd-128">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="f38fd-129">Чтобы не допустить установки надстройки на платформах, которые не поддерживают разметку, добавьте ссылку на набор требований в разделе `<Requirements>` манифест.</span><span class="sxs-lookup"><span data-stu-id="f38fd-129">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="f38fd-130">Инструкции см. в [элементе Set the Requirements in the manifest.](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="f38fd-130">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="f38fd-131">Кроме того, вы можете создать надстройку, чтобы иметь альтернативный опыт, когда **AddinCommands 1.3** не поддерживается, как описано в описании Использования проверок времени запуска в коде [JavaScript.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)</span><span class="sxs-lookup"><span data-stu-id="f38fd-131">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="f38fd-132">Например, если надстройка содержит инструкции, предполагаемые, что встроенные кнопки находятся в настраиваемой группе, вы можете иметь альтернативную версию, предполагаемую, что встроенные кнопки находятся только в обычных местах.</span><span class="sxs-lookup"><span data-stu-id="f38fd-132">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
