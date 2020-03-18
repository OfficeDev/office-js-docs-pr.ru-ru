---
title: Элемент Method в файле манифеста
description: Элемент Method указывает отдельный метод из API JavaScript для Office, необходимый для активации надстроек Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 5da25616d25a8d7454fc847727cda38a9935b5c7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720584"
---
# <a name="method-element"></a><span data-ttu-id="0ae15-103">Элемент Method</span><span class="sxs-lookup"><span data-stu-id="0ae15-103">Method element</span></span>

<span data-ttu-id="0ae15-104">Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="0ae15-104">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="0ae15-105">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="0ae15-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="0ae15-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="0ae15-106">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="0ae15-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="0ae15-107">Contained in</span></span>

[<span data-ttu-id="0ae15-108">Методы</span><span class="sxs-lookup"><span data-stu-id="0ae15-108">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="0ae15-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0ae15-109">Attributes</span></span>

|<span data-ttu-id="0ae15-110">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="0ae15-110">**Attribute**</span></span>|<span data-ttu-id="0ae15-111">**Тип**</span><span class="sxs-lookup"><span data-stu-id="0ae15-111">**Type**</span></span>|<span data-ttu-id="0ae15-112">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="0ae15-112">**Required**</span></span>|<span data-ttu-id="0ae15-113">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0ae15-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0ae15-114">Имя</span><span class="sxs-lookup"><span data-stu-id="0ae15-114">Name</span></span>|<span data-ttu-id="0ae15-115">string</span><span class="sxs-lookup"><span data-stu-id="0ae15-115">string</span></span>|<span data-ttu-id="0ae15-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0ae15-116">required</span></span>|<span data-ttu-id="0ae15-117">Указывает имя необходимого метода, соответствующее его родительскому объекту.</span><span class="sxs-lookup"><span data-stu-id="0ae15-117">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="0ae15-118">Например, чтобы указать `getSelectedDataAsync` метод, необходимо указать. `"Document.getSelectedDataAsync"`</span><span class="sxs-lookup"><span data-stu-id="0ae15-118">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="0ae15-119">Примечания</span><span class="sxs-lookup"><span data-stu-id="0ae15-119">Remarks</span></span>

<span data-ttu-id="0ae15-120">Элементы `Methods` и `Method` не поддерживаются почтовыми надстройками. Дополнительные сведения о наборах требований: [версии и наборы](../../develop/office-versions-and-requirement-sets.md)обязательных элементов для Office.</span><span class="sxs-lookup"><span data-stu-id="0ae15-120">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0ae15-121">Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**.</span><span class="sxs-lookup"><span data-stu-id="0ae15-121">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="0ae15-122">Дополнительные сведения о том, как это сделать, можно узнать в статье Общие сведения об [API JavaScript для Office](../../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="0ae15-122">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
