---
title: Элемент CustomTab в файле манифеста
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: c48e526534a3c1295e9c3f0c6fc626df94a874d3
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554015"
---
# <a name="customtab-element"></a><span data-ttu-id="5078c-102">Элемент CustomTab</span><span class="sxs-lookup"><span data-stu-id="5078c-102">CustomTab element</span></span>

<span data-ttu-id="5078c-p101">На ленте можно указать вкладку и группу для команд надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка.</span><span class="sxs-lookup"><span data-stu-id="5078c-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="5078c-p102">На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="5078c-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="5078c-108">Атрибут **id** должен быть уникальным для манифеста.</span><span class="sxs-lookup"><span data-stu-id="5078c-108">The  **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5078c-109">В Outlook на Mac `CustomTab` элемент недоступен, поэтому необходимо использовать [OfficeTab](officetab.md) .</span><span class="sxs-lookup"><span data-stu-id="5078c-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5078c-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="5078c-110">Child elements</span></span>

|  <span data-ttu-id="5078c-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="5078c-111">Element</span></span> |  <span data-ttu-id="5078c-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5078c-112">Required</span></span>  |  <span data-ttu-id="5078c-113">Описание</span><span class="sxs-lookup"><span data-stu-id="5078c-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5078c-114">Group</span><span class="sxs-lookup"><span data-stu-id="5078c-114">Group</span></span>](group.md)      | <span data-ttu-id="5078c-115">Да</span><span class="sxs-lookup"><span data-stu-id="5078c-115">Yes</span></span> |  <span data-ttu-id="5078c-116">Определяет группу команд.</span><span class="sxs-lookup"><span data-stu-id="5078c-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="5078c-117">Label</span><span class="sxs-lookup"><span data-stu-id="5078c-117">Label</span></span>](#label-tab)      | <span data-ttu-id="5078c-118">Да</span><span class="sxs-lookup"><span data-stu-id="5078c-118">Yes</span></span> |  <span data-ttu-id="5078c-119">Метка элемента CustomTab или Group.</span><span class="sxs-lookup"><span data-stu-id="5078c-119">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="5078c-120">Group</span><span class="sxs-lookup"><span data-stu-id="5078c-120">Group</span></span>

<span data-ttu-id="5078c-p103">Обязательный. См. статью об [элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="5078c-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="5078c-123">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="5078c-123">Label (Tab)</span></span>

<span data-ttu-id="5078c-p104">Обязательный элемент. Метка настраиваемой вкладки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="5078c-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="5078c-126">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="5078c-126">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
