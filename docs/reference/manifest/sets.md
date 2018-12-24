---
title: Элемент Sets в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: b7e78ae05f8409f38c885a1d6a328347d00d0df1
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433658"
---
# <a name="sets-element"></a><span data-ttu-id="f3b55-102">Элемент Sets</span><span class="sxs-lookup"><span data-stu-id="f3b55-102">Sets element</span></span>

<span data-ttu-id="f3b55-103">Указывает минимальное подмножество API JavaScript для Office, необходимое для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="f3b55-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="f3b55-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="f3b55-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f3b55-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="f3b55-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="f3b55-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="f3b55-106">Contained in</span></span>

[<span data-ttu-id="f3b55-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="f3b55-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="f3b55-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="f3b55-108">Can contain</span></span>

[<span data-ttu-id="f3b55-109">Set</span><span class="sxs-lookup"><span data-stu-id="f3b55-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="f3b55-110">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f3b55-110">Attributes</span></span>

|<span data-ttu-id="f3b55-111">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="f3b55-111">**Attribute**</span></span>|<span data-ttu-id="f3b55-112">**Тип**</span><span class="sxs-lookup"><span data-stu-id="f3b55-112">**Type**</span></span>|<span data-ttu-id="f3b55-113">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="f3b55-113">**Required**</span></span>|<span data-ttu-id="f3b55-114">**Описание**</span><span class="sxs-lookup"><span data-stu-id="f3b55-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f3b55-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="f3b55-115">DefaultMinVersion</span></span>|<span data-ttu-id="f3b55-116">string</span><span class="sxs-lookup"><span data-stu-id="f3b55-116">string</span></span>|<span data-ttu-id="f3b55-117">необязательный</span><span class="sxs-lookup"><span data-stu-id="f3b55-117">optional</span></span>|<span data-ttu-id="f3b55-p101">Задает значение атрибута **MinVersion** по умолчанию для всех дочерних элементов [Set](set.md). Значение по умолчанию: "1.1".</span><span class="sxs-lookup"><span data-stu-id="f3b55-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="f3b55-120">Примечания</span><span class="sxs-lookup"><span data-stu-id="f3b55-120">Remarks</span></span>

<span data-ttu-id="f3b55-121">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="f3b55-121">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="f3b55-122">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="f3b55-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

