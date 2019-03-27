---
title: Элемент Sets в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 13777e54ec6bd2d97fa35609ebe194ed85ffa1b8
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871775"
---
# <a name="sets-element"></a><span data-ttu-id="4a844-102">Элемент Sets</span><span class="sxs-lookup"><span data-stu-id="4a844-102">Sets element</span></span>

<span data-ttu-id="4a844-103">Указывает минимальное подмножество API JavaScript для Office, необходимое для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="4a844-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="4a844-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="4a844-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4a844-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="4a844-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="4a844-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="4a844-106">Contained in</span></span>

[<span data-ttu-id="4a844-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a844-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="4a844-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="4a844-108">Can contain</span></span>

[<span data-ttu-id="4a844-109">Set</span><span class="sxs-lookup"><span data-stu-id="4a844-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="4a844-110">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4a844-110">Attributes</span></span>

|<span data-ttu-id="4a844-111">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="4a844-111">**Attribute**</span></span>|<span data-ttu-id="4a844-112">**Тип**</span><span class="sxs-lookup"><span data-stu-id="4a844-112">**Type**</span></span>|<span data-ttu-id="4a844-113">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="4a844-113">**Required**</span></span>|<span data-ttu-id="4a844-114">**Описание**</span><span class="sxs-lookup"><span data-stu-id="4a844-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="4a844-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="4a844-115">DefaultMinVersion</span></span>|<span data-ttu-id="4a844-116">string</span><span class="sxs-lookup"><span data-stu-id="4a844-116">string</span></span>|<span data-ttu-id="4a844-117">необязательный</span><span class="sxs-lookup"><span data-stu-id="4a844-117">optional</span></span>|<span data-ttu-id="4a844-p101">Задает значение атрибута **MinVersion** по умолчанию для всех дочерних элементов [Set](set.md). Значение по умолчанию: "1.1".</span><span class="sxs-lookup"><span data-stu-id="4a844-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="4a844-120">Примечания</span><span class="sxs-lookup"><span data-stu-id="4a844-120">Remarks</span></span>

<span data-ttu-id="4a844-121">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="4a844-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="4a844-122">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="4a844-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

