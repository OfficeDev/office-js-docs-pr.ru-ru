---
title: Элемент Method в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 74b7a8b3d0f8511d21eb0df150500850e8b93fe9
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596895"
---
# <a name="method-element"></a><span data-ttu-id="d3e5f-102">Элемент Method</span><span class="sxs-lookup"><span data-stu-id="d3e5f-102">Method element</span></span>

<span data-ttu-id="d3e5f-103">Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="d3e5f-103">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="d3e5f-104">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="d3e5f-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="d3e5f-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="d3e5f-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="d3e5f-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="d3e5f-106">Contained in</span></span>

[<span data-ttu-id="d3e5f-107">Методы</span><span class="sxs-lookup"><span data-stu-id="d3e5f-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="d3e5f-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d3e5f-108">Attributes</span></span>

|<span data-ttu-id="d3e5f-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="d3e5f-109">**Attribute**</span></span>|<span data-ttu-id="d3e5f-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="d3e5f-110">**Type**</span></span>|<span data-ttu-id="d3e5f-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="d3e5f-111">**Required**</span></span>|<span data-ttu-id="d3e5f-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="d3e5f-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d3e5f-113">Имя</span><span class="sxs-lookup"><span data-stu-id="d3e5f-113">Name</span></span>|<span data-ttu-id="d3e5f-114">string</span><span class="sxs-lookup"><span data-stu-id="d3e5f-114">string</span></span>|<span data-ttu-id="d3e5f-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d3e5f-115">required</span></span>|<span data-ttu-id="d3e5f-116">Указывает имя необходимого метода, соответствующее его родительскому объекту.</span><span class="sxs-lookup"><span data-stu-id="d3e5f-116">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="d3e5f-117">Например, чтобы указать `getSelectedDataAsync` метод, необходимо указать. `"Document.getSelectedDataAsync"`</span><span class="sxs-lookup"><span data-stu-id="d3e5f-117">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="d3e5f-118">Примечания</span><span class="sxs-lookup"><span data-stu-id="d3e5f-118">Remarks</span></span>

<span data-ttu-id="d3e5f-119">Элементы `Methods` и `Method` не поддерживаются почтовыми надстройками. Дополнительные сведения о наборах требований: [версии и наборы](../../develop/office-versions-and-requirement-sets.md)обязательных элементов для Office.</span><span class="sxs-lookup"><span data-stu-id="d3e5f-119">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d3e5f-120">Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**.</span><span class="sxs-lookup"><span data-stu-id="d3e5f-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="d3e5f-121">Дополнительные сведения о том, как это сделать, можно узнать в статье Общие сведения об [API JavaScript для Office](../../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="d3e5f-121">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
