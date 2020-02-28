---
title: Элемент Sets в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 768f674b4afbd65df88825e871005f182d06f6ce
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325243"
---
# <a name="sets-element"></a><span data-ttu-id="bc231-102">Элемент Sets</span><span class="sxs-lookup"><span data-stu-id="bc231-102">Sets element</span></span>

<span data-ttu-id="bc231-103">Указывает минимальное подмножество API JavaScript для Office, необходимое для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="bc231-103">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="bc231-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="bc231-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bc231-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="bc231-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="bc231-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="bc231-106">Contained in</span></span>

[<span data-ttu-id="bc231-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="bc231-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="bc231-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="bc231-108">Can contain</span></span>

[<span data-ttu-id="bc231-109">Set</span><span class="sxs-lookup"><span data-stu-id="bc231-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="bc231-110">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bc231-110">Attributes</span></span>

|<span data-ttu-id="bc231-111">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="bc231-111">**Attribute**</span></span>|<span data-ttu-id="bc231-112">**Тип**</span><span class="sxs-lookup"><span data-stu-id="bc231-112">**Type**</span></span>|<span data-ttu-id="bc231-113">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="bc231-113">**Required**</span></span>|<span data-ttu-id="bc231-114">**Описание**</span><span class="sxs-lookup"><span data-stu-id="bc231-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bc231-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="bc231-115">DefaultMinVersion</span></span>|<span data-ttu-id="bc231-116">string</span><span class="sxs-lookup"><span data-stu-id="bc231-116">string</span></span>|<span data-ttu-id="bc231-117">необязательный</span><span class="sxs-lookup"><span data-stu-id="bc231-117">optional</span></span>|<span data-ttu-id="bc231-118">Задает значение атрибута **MinVersion** по умолчанию для всех дочерних элементов [набора](set.md) .</span><span class="sxs-lookup"><span data-stu-id="bc231-118">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="bc231-119">Значение по умолчанию: "1.1".</span><span class="sxs-lookup"><span data-stu-id="bc231-119">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="bc231-120">Примечания</span><span class="sxs-lookup"><span data-stu-id="bc231-120">Remarks</span></span>

<span data-ttu-id="bc231-121">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="bc231-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="bc231-122">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **дефаултминверсион** элемента **Sets** приведены в разделе [set the требований в манифесте](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="bc231-122">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

