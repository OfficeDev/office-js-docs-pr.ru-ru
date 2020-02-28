---
title: Элемент CustomTab в файле манифеста
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: ba0419b6cf9cc4a0c1e3038dbb7f972e65868ec4
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42323807"
---
# <a name="customtab-element"></a><span data-ttu-id="9edbf-102">Элемент CustomTab</span><span class="sxs-lookup"><span data-stu-id="9edbf-102">CustomTab element</span></span>

<span data-ttu-id="9edbf-103">На ленте можно указать вкладку и группу для команд надстройки.</span><span class="sxs-lookup"><span data-stu-id="9edbf-103">On the ribbon, you specify which tab and group for their add-in commands.</span></span> <span data-ttu-id="9edbf-104">Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="9edbf-104">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="9edbf-p102">На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="9edbf-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="9edbf-108">Атрибут **ID** должен быть уникальным в пределах манифеста.</span><span class="sxs-lookup"><span data-stu-id="9edbf-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9edbf-109">В Outlook на Mac `CustomTab` элемент недоступен, поэтому необходимо использовать [OfficeTab](officetab.md) .</span><span class="sxs-lookup"><span data-stu-id="9edbf-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9edbf-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9edbf-110">Child elements</span></span>

|  <span data-ttu-id="9edbf-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="9edbf-111">Element</span></span> |  <span data-ttu-id="9edbf-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9edbf-112">Required</span></span>  |  <span data-ttu-id="9edbf-113">Описание</span><span class="sxs-lookup"><span data-stu-id="9edbf-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9edbf-114">Group</span><span class="sxs-lookup"><span data-stu-id="9edbf-114">Group</span></span>](group.md)      | <span data-ttu-id="9edbf-115">Да</span><span class="sxs-lookup"><span data-stu-id="9edbf-115">Yes</span></span> |  <span data-ttu-id="9edbf-116">Определяет группу команд.</span><span class="sxs-lookup"><span data-stu-id="9edbf-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="9edbf-117">Label</span><span class="sxs-lookup"><span data-stu-id="9edbf-117">Label</span></span>](#label-tab)      | <span data-ttu-id="9edbf-118">Да</span><span class="sxs-lookup"><span data-stu-id="9edbf-118">Yes</span></span> |  <span data-ttu-id="9edbf-119">Метка элемента CustomTab или Group.</span><span class="sxs-lookup"><span data-stu-id="9edbf-119">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="9edbf-120">Group</span><span class="sxs-lookup"><span data-stu-id="9edbf-120">Group</span></span>

<span data-ttu-id="9edbf-p103">Обязательный. См. статью об [элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="9edbf-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="9edbf-123">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="9edbf-123">Label (Tab)</span></span>

<span data-ttu-id="9edbf-124">Обязательное.</span><span class="sxs-lookup"><span data-stu-id="9edbf-124">Required.</span></span> <span data-ttu-id="9edbf-125">Метка настраиваемой вкладки. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="9edbf-125">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="9edbf-126">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="9edbf-126">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
