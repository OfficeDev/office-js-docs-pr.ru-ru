---
title: Элемент Set в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 47f675f999a225e499171cb03c27797bb3dcc5f6
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596510"
---
# <a name="set-element"></a><span data-ttu-id="59511-102">Элемент Set</span><span class="sxs-lookup"><span data-stu-id="59511-102">Set element</span></span>

<span data-ttu-id="59511-103">Задает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="59511-103">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="59511-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="59511-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="59511-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="59511-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="59511-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="59511-106">Contained in</span></span>

[<span data-ttu-id="59511-107">Sets</span><span class="sxs-lookup"><span data-stu-id="59511-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="59511-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="59511-108">Attributes</span></span>

|<span data-ttu-id="59511-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="59511-109">**Attribute**</span></span>|<span data-ttu-id="59511-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="59511-110">**Type**</span></span>|<span data-ttu-id="59511-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="59511-111">**Required**</span></span>|<span data-ttu-id="59511-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="59511-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="59511-113">Имя</span><span class="sxs-lookup"><span data-stu-id="59511-113">Name</span></span>|<span data-ttu-id="59511-114">string</span><span class="sxs-lookup"><span data-stu-id="59511-114">string</span></span>|<span data-ttu-id="59511-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="59511-115">required</span></span>|<span data-ttu-id="59511-116">Имя [набора требований](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="59511-116">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="59511-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="59511-117">MinVersion</span></span>|<span data-ttu-id="59511-118">string</span><span class="sxs-lookup"><span data-stu-id="59511-118">string</span></span>|<span data-ttu-id="59511-119">необязательный</span><span class="sxs-lookup"><span data-stu-id="59511-119">optional</span></span>|<span data-ttu-id="59511-120">Указывает минимальную версию набора API, необходимую надстройке.</span><span class="sxs-lookup"><span data-stu-id="59511-120">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="59511-121">Переопределяет значение **дефаултминверсион**, если оно указано в элементе родительских [наборов](sets.md) .</span><span class="sxs-lookup"><span data-stu-id="59511-121">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="59511-122">Примечания</span><span class="sxs-lookup"><span data-stu-id="59511-122">Remarks</span></span>

<span data-ttu-id="59511-123">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="59511-123">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="59511-124">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **дефаултминверсион** элемента **Sets** приведены в разделе [set the требований в манифесте](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="59511-124">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="59511-125">Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`.</span><span class="sxs-lookup"><span data-stu-id="59511-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="59511-126">Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек).</span><span class="sxs-lookup"><span data-stu-id="59511-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="59511-127">Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.</span><span class="sxs-lookup"><span data-stu-id="59511-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
