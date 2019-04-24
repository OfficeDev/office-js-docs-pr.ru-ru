---
title: Элемент Set в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0f408d698d297eaa6287ff268bdb7fc737a5a24d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452034"
---
# <a name="set-element"></a><span data-ttu-id="24c75-102">Элемент Set</span><span class="sxs-lookup"><span data-stu-id="24c75-102">Set element</span></span>

<span data-ttu-id="24c75-103">Указывает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="24c75-103">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="24c75-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="24c75-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="24c75-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="24c75-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="24c75-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="24c75-106">Contained in</span></span>

[<span data-ttu-id="24c75-107">Sets</span><span class="sxs-lookup"><span data-stu-id="24c75-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="24c75-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="24c75-108">Attributes</span></span>

|<span data-ttu-id="24c75-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="24c75-109">**Attribute**</span></span>|<span data-ttu-id="24c75-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="24c75-110">**Type**</span></span>|<span data-ttu-id="24c75-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="24c75-111">**Required**</span></span>|<span data-ttu-id="24c75-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="24c75-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="24c75-113">Имя</span><span class="sxs-lookup"><span data-stu-id="24c75-113">Name</span></span>|<span data-ttu-id="24c75-114">строка</span><span class="sxs-lookup"><span data-stu-id="24c75-114">string</span></span>|<span data-ttu-id="24c75-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="24c75-115">required</span></span>|<span data-ttu-id="24c75-116">Имя [набора требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="24c75-116">The name of a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="24c75-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="24c75-117">MinVersion</span></span>|<span data-ttu-id="24c75-118">string</span><span class="sxs-lookup"><span data-stu-id="24c75-118">string</span></span>|<span data-ttu-id="24c75-119">необязательный</span><span class="sxs-lookup"><span data-stu-id="24c75-119">optional</span></span>|<span data-ttu-id="24c75-p101">Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion**, если оно указано в родительском элементе [Sets](sets.md).</span><span class="sxs-lookup"><span data-stu-id="24c75-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="24c75-122">Примечания</span><span class="sxs-lookup"><span data-stu-id="24c75-122">Remarks</span></span>

<span data-ttu-id="24c75-123">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="24c75-123">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="24c75-124">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="24c75-124">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="24c75-125">Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`.</span><span class="sxs-lookup"><span data-stu-id="24c75-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="24c75-126">Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек).</span><span class="sxs-lookup"><span data-stu-id="24c75-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="24c75-127">Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.</span><span class="sxs-lookup"><span data-stu-id="24c75-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
