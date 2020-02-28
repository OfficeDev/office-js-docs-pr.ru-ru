---
title: Элемент Supertip в файле манифеста
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: ab280ec550a58f85082c36a24f5f7c3b4112a214
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325236"
---
# <a name="supertip"></a><span data-ttu-id="dd63c-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="dd63c-102">Supertip</span></span>

<span data-ttu-id="dd63c-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="dd63c-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="dd63c-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="dd63c-105">Child elements</span></span>

|  <span data-ttu-id="dd63c-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd63c-106">Element</span></span> |  <span data-ttu-id="dd63c-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="dd63c-107">Required</span></span>  |  <span data-ttu-id="dd63c-108">Описание</span><span class="sxs-lookup"><span data-stu-id="dd63c-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="dd63c-109">Title</span><span class="sxs-lookup"><span data-stu-id="dd63c-109">Title</span></span>](#title) | <span data-ttu-id="dd63c-110">Да</span><span class="sxs-lookup"><span data-stu-id="dd63c-110">Yes</span></span> | <span data-ttu-id="dd63c-111">Текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="dd63c-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="dd63c-112">Description</span><span class="sxs-lookup"><span data-stu-id="dd63c-112">Description</span></span>](#description) | <span data-ttu-id="dd63c-113">Да</span><span class="sxs-lookup"><span data-stu-id="dd63c-113">Yes</span></span> | <span data-ttu-id="dd63c-114">Описание подсказки.</span><span class="sxs-lookup"><span data-stu-id="dd63c-114">The description for the supertip.</span></span><br><span data-ttu-id="dd63c-115">**Note**: (Outlook) поддерживаются только клиенты Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="dd63c-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="dd63c-116">Название</span><span class="sxs-lookup"><span data-stu-id="dd63c-116">Title</span></span>

<span data-ttu-id="dd63c-117">Обязательное.</span><span class="sxs-lookup"><span data-stu-id="dd63c-117">Required.</span></span> <span data-ttu-id="dd63c-118">Текст суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="dd63c-118">The text for the supertip.</span></span> <span data-ttu-id="dd63c-119">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="dd63c-119">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="dd63c-120">Описание</span><span class="sxs-lookup"><span data-stu-id="dd63c-120">Description</span></span>

<span data-ttu-id="dd63c-121">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="dd63c-121">Required.</span></span> <span data-ttu-id="dd63c-122">Описание суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="dd63c-122">The description for the supertip.</span></span> <span data-ttu-id="dd63c-123">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **LongStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="dd63c-123">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="dd63c-124">В Outlook только клиенты Windows и Mac поддерживают элемент **Description** .</span><span class="sxs-lookup"><span data-stu-id="dd63c-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="dd63c-125">Пример</span><span class="sxs-lookup"><span data-stu-id="dd63c-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
