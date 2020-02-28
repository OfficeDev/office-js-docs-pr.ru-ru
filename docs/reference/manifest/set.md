---
title: Элемент Set в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d86b3123ff856e8618f92629308787b543e8228b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324808"
---
# <a name="set-element"></a><span data-ttu-id="2dea5-102">Элемент Set</span><span class="sxs-lookup"><span data-stu-id="2dea5-102">Set element</span></span>

<span data-ttu-id="2dea5-103">Задает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="2dea5-103">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="2dea5-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="2dea5-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2dea5-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2dea5-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="2dea5-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2dea5-106">Contained in</span></span>

[<span data-ttu-id="2dea5-107">Sets</span><span class="sxs-lookup"><span data-stu-id="2dea5-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="2dea5-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2dea5-108">Attributes</span></span>

|<span data-ttu-id="2dea5-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="2dea5-109">**Attribute**</span></span>|<span data-ttu-id="2dea5-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="2dea5-110">**Type**</span></span>|<span data-ttu-id="2dea5-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="2dea5-111">**Required**</span></span>|<span data-ttu-id="2dea5-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="2dea5-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2dea5-113">Имя</span><span class="sxs-lookup"><span data-stu-id="2dea5-113">Name</span></span>|<span data-ttu-id="2dea5-114">string</span><span class="sxs-lookup"><span data-stu-id="2dea5-114">string</span></span>|<span data-ttu-id="2dea5-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2dea5-115">required</span></span>|<span data-ttu-id="2dea5-116">Имя [набора требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2dea5-116">The name of a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="2dea5-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="2dea5-117">MinVersion</span></span>|<span data-ttu-id="2dea5-118">string</span><span class="sxs-lookup"><span data-stu-id="2dea5-118">string</span></span>|<span data-ttu-id="2dea5-119">необязательный</span><span class="sxs-lookup"><span data-stu-id="2dea5-119">optional</span></span>|<span data-ttu-id="2dea5-120">Указывает минимальную версию набора API, необходимую надстройке.</span><span class="sxs-lookup"><span data-stu-id="2dea5-120">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="2dea5-121">Переопределяет значение **дефаултминверсион**, если оно указано в элементе родительских [наборов](sets.md) .</span><span class="sxs-lookup"><span data-stu-id="2dea5-121">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="2dea5-122">Примечания</span><span class="sxs-lookup"><span data-stu-id="2dea5-122">Remarks</span></span>

<span data-ttu-id="2dea5-123">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2dea5-123">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="2dea5-124">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **дефаултминверсион** элемента **Sets** приведены в разделе [set the требований в манифесте](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="2dea5-124">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="2dea5-125">Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`.</span><span class="sxs-lookup"><span data-stu-id="2dea5-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="2dea5-126">Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек).</span><span class="sxs-lookup"><span data-stu-id="2dea5-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="2dea5-127">Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.</span><span class="sxs-lookup"><span data-stu-id="2dea5-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
