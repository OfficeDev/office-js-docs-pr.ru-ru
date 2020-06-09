---
title: Элемент Set в файле манифеста
description: Элемент Set указывает набор обязательных элементов API JavaScript для Office, необходимый для активации надстройки Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f4755cc6742beb37ed8b8efcf4c3968394f15ed6
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608742"
---
# <a name="set-element"></a><span data-ttu-id="b72ce-103">Элемент Set</span><span class="sxs-lookup"><span data-stu-id="b72ce-103">Set element</span></span>

<span data-ttu-id="b72ce-104">Задает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="b72ce-104">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="b72ce-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="b72ce-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b72ce-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b72ce-106">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="b72ce-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="b72ce-107">Contained in</span></span>

[<span data-ttu-id="b72ce-108">Sets</span><span class="sxs-lookup"><span data-stu-id="b72ce-108">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="b72ce-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b72ce-109">Attributes</span></span>

|<span data-ttu-id="b72ce-110">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="b72ce-110">**Attribute**</span></span>|<span data-ttu-id="b72ce-111">**Тип**</span><span class="sxs-lookup"><span data-stu-id="b72ce-111">**Type**</span></span>|<span data-ttu-id="b72ce-112">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="b72ce-112">**Required**</span></span>|<span data-ttu-id="b72ce-113">**Описание**</span><span class="sxs-lookup"><span data-stu-id="b72ce-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="b72ce-114">Имя</span><span class="sxs-lookup"><span data-stu-id="b72ce-114">Name</span></span>|<span data-ttu-id="b72ce-115">string</span><span class="sxs-lookup"><span data-stu-id="b72ce-115">string</span></span>|<span data-ttu-id="b72ce-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="b72ce-116">required</span></span>|<span data-ttu-id="b72ce-117">Имя [набора требований](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b72ce-117">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="b72ce-118">MinVersion</span><span class="sxs-lookup"><span data-stu-id="b72ce-118">MinVersion</span></span>|<span data-ttu-id="b72ce-119">string</span><span class="sxs-lookup"><span data-stu-id="b72ce-119">string</span></span>|<span data-ttu-id="b72ce-120">необязательный</span><span class="sxs-lookup"><span data-stu-id="b72ce-120">optional</span></span>|<span data-ttu-id="b72ce-121">Указывает минимальную версию набора API, необходимую надстройке.</span><span class="sxs-lookup"><span data-stu-id="b72ce-121">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="b72ce-122">Переопределяет значение **дефаултминверсион**, если оно указано в элементе родительских [наборов](sets.md) .</span><span class="sxs-lookup"><span data-stu-id="b72ce-122">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="b72ce-123">Примечания</span><span class="sxs-lookup"><span data-stu-id="b72ce-123">Remarks</span></span>

<span data-ttu-id="b72ce-124">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b72ce-124">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="b72ce-125">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **дефаултминверсион** элемента **Sets** приведены в разделе [set the требований в манифесте](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="b72ce-125">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="b72ce-126">Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`.</span><span class="sxs-lookup"><span data-stu-id="b72ce-126">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="b72ce-127">Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек).</span><span class="sxs-lookup"><span data-stu-id="b72ce-127">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="b72ce-128">Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.</span><span class="sxs-lookup"><span data-stu-id="b72ce-128">Also, you can't declare support for specific methods in mail add-ins.</span></span>
