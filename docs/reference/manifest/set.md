---
title: Элемент Set в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0f137f7b08d6f1d0b0d972173c8085713b0f979d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432769"
---
# <a name="set-element"></a><span data-ttu-id="2238a-102">Элемент Set</span><span class="sxs-lookup"><span data-stu-id="2238a-102">Set element</span></span>

<span data-ttu-id="2238a-103">Указывает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="2238a-103">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="2238a-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="2238a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2238a-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2238a-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="2238a-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2238a-106">Contained in</span></span>

[<span data-ttu-id="2238a-107">Sets</span><span class="sxs-lookup"><span data-stu-id="2238a-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="2238a-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2238a-108">Attributes</span></span>

|<span data-ttu-id="2238a-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="2238a-109">**Attribute**</span></span>|<span data-ttu-id="2238a-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="2238a-110">**Type**</span></span>|<span data-ttu-id="2238a-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="2238a-111">**Required**</span></span>|<span data-ttu-id="2238a-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="2238a-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2238a-113">Имя</span><span class="sxs-lookup"><span data-stu-id="2238a-113">Name</span></span>|<span data-ttu-id="2238a-114">string</span><span class="sxs-lookup"><span data-stu-id="2238a-114">string</span></span>|<span data-ttu-id="2238a-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2238a-115">required</span></span>|<span data-ttu-id="2238a-116">Имя [набора требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2238a-116">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="2238a-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="2238a-117">MinVersion</span></span>|<span data-ttu-id="2238a-118">string</span><span class="sxs-lookup"><span data-stu-id="2238a-118">string</span></span>|<span data-ttu-id="2238a-119">необязательный</span><span class="sxs-lookup"><span data-stu-id="2238a-119">optional</span></span>|<span data-ttu-id="2238a-p101">Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion**, если оно указано в родительском элементе [Sets](sets.md).</span><span class="sxs-lookup"><span data-stu-id="2238a-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="2238a-122">Примечания</span><span class="sxs-lookup"><span data-stu-id="2238a-122">Remarks</span></span>

<span data-ttu-id="2238a-123">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2238a-123">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="2238a-124">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="2238a-124">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="2238a-125">Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`.</span><span class="sxs-lookup"><span data-stu-id="2238a-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="2238a-126">Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек).</span><span class="sxs-lookup"><span data-stu-id="2238a-126">Important  For mail add-ins, there is only one   requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.</span></span> <span data-ttu-id="2238a-127">Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.</span><span class="sxs-lookup"><span data-stu-id="2238a-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
