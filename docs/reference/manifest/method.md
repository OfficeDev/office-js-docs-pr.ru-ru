---
title: Элемент Method в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2bcc24abf269f5d6c44c03e738bac480fd05d5ca
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324850"
---
# <a name="method-element"></a><span data-ttu-id="9a136-102">Элемент Method</span><span class="sxs-lookup"><span data-stu-id="9a136-102">Method element</span></span>

<span data-ttu-id="9a136-103">Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9a136-103">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="9a136-104">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="9a136-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="9a136-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="9a136-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="9a136-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="9a136-106">Contained in</span></span>

[<span data-ttu-id="9a136-107">Методы</span><span class="sxs-lookup"><span data-stu-id="9a136-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="9a136-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9a136-108">Attributes</span></span>

|<span data-ttu-id="9a136-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="9a136-109">**Attribute**</span></span>|<span data-ttu-id="9a136-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="9a136-110">**Type**</span></span>|<span data-ttu-id="9a136-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="9a136-111">**Required**</span></span>|<span data-ttu-id="9a136-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="9a136-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="9a136-113">Имя</span><span class="sxs-lookup"><span data-stu-id="9a136-113">Name</span></span>|<span data-ttu-id="9a136-114">string</span><span class="sxs-lookup"><span data-stu-id="9a136-114">string</span></span>|<span data-ttu-id="9a136-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9a136-115">required</span></span>|<span data-ttu-id="9a136-116">Указывает имя необходимого метода, соответствующее его родительскому объекту.</span><span class="sxs-lookup"><span data-stu-id="9a136-116">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="9a136-117">Например, чтобы указать `getSelectedDataAsync` метод, необходимо указать. `"Document.getSelectedDataAsync"`</span><span class="sxs-lookup"><span data-stu-id="9a136-117">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="9a136-118">Замечания</span><span class="sxs-lookup"><span data-stu-id="9a136-118">Remarks</span></span>

<span data-ttu-id="9a136-119">Элементы `Methods` и `Method` не поддерживаются почтовыми надстройками. Дополнительные сведения о наборах требований: [версии и наборы](/office/dev/add-ins/develop/office-versions-and-requirement-sets)обязательных элементов для Office.</span><span class="sxs-lookup"><span data-stu-id="9a136-119">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="9a136-120">Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**.</span><span class="sxs-lookup"><span data-stu-id="9a136-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="9a136-121">Дополнительные сведения о том, как это сделать, можно узнать в статье Общие сведения об [API JavaScript для Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="9a136-121">For more information about how to do this, see [Understanding the Office JavaScript API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

