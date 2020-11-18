---
title: Интеграция встроенных кнопок Office в группы и вкладки пользовательских элементов управления
description: Узнайте, как включать встроенные кнопки Office в настраиваемые группы команд и вкладки ленты Office.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: e04107893b3c0dd453c84d38fdd5623e308b70e3
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088177"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a><span data-ttu-id="732e1-103">Интеграция встроенных кнопок Office в группы и вкладки пользовательских элементов управления (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="732e1-103">Integrate built-in Office buttons into custom control groups and tabs (preview)</span></span>

<span data-ttu-id="732e1-104">Встроенные кнопки Office можно вставлять в пользовательские группы элементов управления на ленте Office с помощью разметки в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="732e1-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="732e1-105">(Вы не можете вставить свои команды надстройки во встроенную группу Office.) Вы также можете вставить все встроенные группы элементов управления Office в пользовательские вкладки ленты.</span><span class="sxs-lookup"><span data-stu-id="732e1-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="732e1-106">В этой статье предполагается, что вы знакомы с [основными понятиями, изложенными в](add-in-commands.md)статьях, посвященных командам надстроек.</span><span class="sxs-lookup"><span data-stu-id="732e1-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="732e1-107">Изучите его, если это не было сделано недавно.</span><span class="sxs-lookup"><span data-stu-id="732e1-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="732e1-108">Функция надстройки и разметка, описанные в этой статье, *доступны только в PowerPoint в Интернете*.</span><span class="sxs-lookup"><span data-stu-id="732e1-108">The add-in feature and markup described in this article is in preview and is *only available in PowerPoint on the web*.</span></span> <span data-ttu-id="732e1-109">Мы рекомендуем испытать разметку только в средах тестирования и разработки.</span><span class="sxs-lookup"><span data-stu-id="732e1-109">We recommend that you try out the markup in test and development environments only.</span></span> <span data-ttu-id="732e1-110">Не используйте разметку предварительного просмотра в рабочей среде или в критически важных для бизнеса документах.</span><span class="sxs-lookup"><span data-stu-id="732e1-110">Do not use preview markup in a production environment or within business-critical documents.</span></span>
> - <span data-ttu-id="732e1-111">Разметка, описанная в этой статье, работает только на платформах, поддерживающих набор требований **аддинкоммандс 1,3**.</span><span class="sxs-lookup"><span data-stu-id="732e1-111">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="732e1-112">Ознакомьтесь с разделом [поведение неподдерживаемых платформ](#behavior-on-unsupported-platforms).</span><span class="sxs-lookup"><span data-stu-id="732e1-112">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="732e1-113">Вставка встроенной группы управления на настраиваемую вкладку</span><span class="sxs-lookup"><span data-stu-id="732e1-113">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="732e1-114">Чтобы вставить встроенную группу управления Office на вкладку, добавьте элемент [оффицеграуп](../reference/manifest/customtab.md#officegroup) в качестве дочернего элемента в родительском `<CustomTab>` элементе.</span><span class="sxs-lookup"><span data-stu-id="732e1-114">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="732e1-115">`id` `<OfficeGroup>` Для атрибута элемента задается идентификатор встроенной группы.</span><span class="sxs-lookup"><span data-stu-id="732e1-115">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="732e1-116">Узнайте [, как найти идентификаторы элементов управления и групп элементов управления](#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="732e1-116">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="732e1-117">В следующем примере разметки группа элементов управления абзаца Office добавляется на настраиваемую вкладку и помещается сразу после настраиваемой группы.</span><span class="sxs-lookup"><span data-stu-id="732e1-117">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="732e1-118">Вставка встроенного элемента управления в пользовательскую группу</span><span class="sxs-lookup"><span data-stu-id="732e1-118">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="732e1-119">Чтобы вставить встроенный элемент управления Office в настраиваемую группу, добавьте элемент [оффицеконтрол](../reference/manifest/group.md#officecontrol) в качестве дочернего элемента в родительском `<Group>` элементе.</span><span class="sxs-lookup"><span data-stu-id="732e1-119">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="732e1-120">`id` `<OfficeControl>` Для атрибута элемента задается Идентификатор встроенного элемента управления.</span><span class="sxs-lookup"><span data-stu-id="732e1-120">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="732e1-121">Узнайте [, как найти идентификаторы элементов управления и групп элементов управления](#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="732e1-121">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="732e1-122">В приведенном ниже примере разметки элемент управления "верхний скрипт Office" добавляется в пользовательскую группу и помещается сразу после настраиваемой кнопки.</span><span class="sxs-lookup"><span data-stu-id="732e1-122">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

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
> <span data-ttu-id="732e1-123">Пользователи могут настраивать ленту в приложении Office.</span><span class="sxs-lookup"><span data-stu-id="732e1-123">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="732e1-124">Все пользовательские настройки будут переопределять параметры манифеста.</span><span class="sxs-lookup"><span data-stu-id="732e1-124">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="732e1-125">Например, пользователь может удалить кнопку из любой группы и удалить любую группу из вкладки.</span><span class="sxs-lookup"><span data-stu-id="732e1-125">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="732e1-126">Поиск идентификаторов элементов управления и групп элементов управления</span><span class="sxs-lookup"><span data-stu-id="732e1-126">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="732e1-127">Идентификаторы для поддерживаемых элементов управления и групп элементов управления находятся в файлах в [идентификаторах элементов управления для Организации](https://github.com/OfficeDev/office-control-ids)в репозитории.</span><span class="sxs-lookup"><span data-stu-id="732e1-127">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="732e1-128">Следуйте инструкциям в файле ReadMe в этом репозитории.</span><span class="sxs-lookup"><span data-stu-id="732e1-128">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="732e1-129">Поведение на неподдерживаемых платформах</span><span class="sxs-lookup"><span data-stu-id="732e1-129">Behavior on unsupported platforms</span></span>

<span data-ttu-id="732e1-130">Если ваша надстройка установлена на платформе, которая не поддерживает [набор требований аддинкоммандс 1,3](../reference/requirement-sets/add-in-commands-requirement-sets.md), разметка, описанная в этой статье, игнорируется, и встроенные элементы управления и группы Office не будут отображаться в настраиваемых группах и вкладках.</span><span class="sxs-lookup"><span data-stu-id="732e1-130">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="732e1-131">Чтобы запретить установку надстройки на платформах, которые не поддерживают разметку, добавьте ссылку на набор требований в `<Requirements>` разделе манифеста.</span><span class="sxs-lookup"><span data-stu-id="732e1-131">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="732e1-132">Инструкции см в разделе [set the требований element в манифесте](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="732e1-132">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="732e1-133">Кроме того, вы можете создать альтернативный интерфейс для надстройки, когда **аддинкоммандс 1,3** не поддерживается, как описано в статье [Использование проверок среды выполнения в коде JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="732e1-133">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="732e1-134">Например, если ваша надстройка содержит инструкции, которые предполагают, что встроенные кнопки находятся в настраиваемых группах, можно использовать альтернативную версию, которая предполагает, что встроенные кнопки находятся только в их обычных местах.</span><span class="sxs-lookup"><span data-stu-id="732e1-134">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
