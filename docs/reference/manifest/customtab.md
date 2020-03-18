---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8db29f166b5a5238a7ecf121ba5e5adca66ebe94
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718428"
---
# <a name="customtab-element"></a><span data-ttu-id="ebd98-103">Элемент CustomTab</span><span class="sxs-lookup"><span data-stu-id="ebd98-103">CustomTab element</span></span>

<span data-ttu-id="ebd98-104">На ленте можно указать вкладку и группу для команд надстройки.</span><span class="sxs-lookup"><span data-stu-id="ebd98-104">On the ribbon, you specify which tab and group for their add-in commands.</span></span> <span data-ttu-id="ebd98-105">Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="ebd98-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="ebd98-p102">На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="ebd98-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="ebd98-109">Атрибут **ID** должен быть уникальным в пределах манифеста.</span><span class="sxs-lookup"><span data-stu-id="ebd98-109">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ebd98-110">В Outlook на Mac `CustomTab` элемент недоступен, поэтому необходимо использовать [OfficeTab](officetab.md) .</span><span class="sxs-lookup"><span data-stu-id="ebd98-110">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ebd98-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ebd98-111">Child elements</span></span>

|  <span data-ttu-id="ebd98-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="ebd98-112">Element</span></span> |  <span data-ttu-id="ebd98-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ebd98-113">Required</span></span>  |  <span data-ttu-id="ebd98-114">Описание</span><span class="sxs-lookup"><span data-stu-id="ebd98-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ebd98-115">Group</span><span class="sxs-lookup"><span data-stu-id="ebd98-115">Group</span></span>](group.md)      | <span data-ttu-id="ebd98-116">Да</span><span class="sxs-lookup"><span data-stu-id="ebd98-116">Yes</span></span> |  <span data-ttu-id="ebd98-117">Определяет группу команд.</span><span class="sxs-lookup"><span data-stu-id="ebd98-117">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="ebd98-118">Label</span><span class="sxs-lookup"><span data-stu-id="ebd98-118">Label</span></span>](#label-tab)      | <span data-ttu-id="ebd98-119">Да</span><span class="sxs-lookup"><span data-stu-id="ebd98-119">Yes</span></span> |  <span data-ttu-id="ebd98-120">Метка элемента CustomTab или Group.</span><span class="sxs-lookup"><span data-stu-id="ebd98-120">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="ebd98-121">Group</span><span class="sxs-lookup"><span data-stu-id="ebd98-121">Group</span></span>

<span data-ttu-id="ebd98-p103">Обязательный. См. статью об [элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="ebd98-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="ebd98-124">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="ebd98-124">Label (Tab)</span></span>

<span data-ttu-id="ebd98-125">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="ebd98-125">Required.</span></span> <span data-ttu-id="ebd98-126">Метка настраиваемой вкладки. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="ebd98-126">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="ebd98-127">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="ebd98-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
