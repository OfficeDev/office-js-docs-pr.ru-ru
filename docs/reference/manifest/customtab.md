---
title: Элемент CustomTab в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 7d609ad216ba5e8e7358bbc741f7b6c992bc97e2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433609"
---
# <a name="customtab-element"></a><span data-ttu-id="718d3-102">Элемент CustomTab</span><span class="sxs-lookup"><span data-stu-id="718d3-102">CustomTab element</span></span>

<span data-ttu-id="718d3-p101">На ленте можно указать вкладку и группу для команд надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка.</span><span class="sxs-lookup"><span data-stu-id="718d3-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="718d3-p102">На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="718d3-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="718d3-108">Атрибут **id** должен быть уникальным для манифеста.</span><span class="sxs-lookup"><span data-stu-id="718d3-108">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="718d3-109">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="718d3-109">Child elements</span></span>

|  <span data-ttu-id="718d3-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="718d3-110">Element</span></span> |  <span data-ttu-id="718d3-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="718d3-111">Required</span></span>  |  <span data-ttu-id="718d3-112">Описание</span><span class="sxs-lookup"><span data-stu-id="718d3-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="718d3-113">Group</span><span class="sxs-lookup"><span data-stu-id="718d3-113">Group</span></span>](group.md)      | <span data-ttu-id="718d3-114">Да</span><span class="sxs-lookup"><span data-stu-id="718d3-114">Yes</span></span> |  <span data-ttu-id="718d3-115">Определяет группу команд.</span><span class="sxs-lookup"><span data-stu-id="718d3-115">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="718d3-116">Label</span><span class="sxs-lookup"><span data-stu-id="718d3-116">Label</span></span>](#label-tab)      | <span data-ttu-id="718d3-117">Да</span><span class="sxs-lookup"><span data-stu-id="718d3-117">Yes</span></span> |  <span data-ttu-id="718d3-118">Метка элемента CustomTab или Group.</span><span class="sxs-lookup"><span data-stu-id="718d3-118">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="718d3-119">Control</span><span class="sxs-lookup"><span data-stu-id="718d3-119">Control</span></span>](control.md)    | <span data-ttu-id="718d3-120">Да</span><span class="sxs-lookup"><span data-stu-id="718d3-120">Yes</span></span> |  <span data-ttu-id="718d3-121">Коллекция из одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="718d3-121">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="718d3-122">Group</span><span class="sxs-lookup"><span data-stu-id="718d3-122">Group</span></span>

<span data-ttu-id="718d3-p103">Обязательный. См. статью об [элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="718d3-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="718d3-125">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="718d3-125">Label (Tab)</span></span>

<span data-ttu-id="718d3-p104">Обязательный элемент. Метка настраиваемой вкладки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="718d3-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="718d3-128">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="718d3-128">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```