---
title: Элемент Sets в файле манифеста
description: Элемент Sets указывает минимальный набор API JavaScript для Office, необходимый для активации надстройки Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c9e699b4609004c49d954da2367a6c8f82d13670
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720395"
---
# <a name="sets-element"></a><span data-ttu-id="511ce-103">Элемент Sets</span><span class="sxs-lookup"><span data-stu-id="511ce-103">Sets element</span></span>

<span data-ttu-id="511ce-104">Указывает минимальное подмножество API JavaScript для Office, необходимое для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="511ce-104">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="511ce-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="511ce-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="511ce-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="511ce-106">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="511ce-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="511ce-107">Contained in</span></span>

[<span data-ttu-id="511ce-108">Requirements</span><span class="sxs-lookup"><span data-stu-id="511ce-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="511ce-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="511ce-109">Can contain</span></span>

[<span data-ttu-id="511ce-110">Set</span><span class="sxs-lookup"><span data-stu-id="511ce-110">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="511ce-111">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="511ce-111">Attributes</span></span>

|<span data-ttu-id="511ce-112">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="511ce-112">**Attribute**</span></span>|<span data-ttu-id="511ce-113">**Тип**</span><span class="sxs-lookup"><span data-stu-id="511ce-113">**Type**</span></span>|<span data-ttu-id="511ce-114">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="511ce-114">**Required**</span></span>|<span data-ttu-id="511ce-115">**Описание**</span><span class="sxs-lookup"><span data-stu-id="511ce-115">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="511ce-116">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="511ce-116">DefaultMinVersion</span></span>|<span data-ttu-id="511ce-117">string</span><span class="sxs-lookup"><span data-stu-id="511ce-117">string</span></span>|<span data-ttu-id="511ce-118">необязательный</span><span class="sxs-lookup"><span data-stu-id="511ce-118">optional</span></span>|<span data-ttu-id="511ce-119">Задает значение атрибута **MinVersion** по умолчанию для всех дочерних элементов [набора](set.md) .</span><span class="sxs-lookup"><span data-stu-id="511ce-119">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="511ce-120">Значение по умолчанию: "1.1".</span><span class="sxs-lookup"><span data-stu-id="511ce-120">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="511ce-121">Примечания</span><span class="sxs-lookup"><span data-stu-id="511ce-121">Remarks</span></span>

<span data-ttu-id="511ce-122">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="511ce-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="511ce-123">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **дефаултминверсион** элемента **Sets** приведены в разделе [set the требований в манифесте](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="511ce-123">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

