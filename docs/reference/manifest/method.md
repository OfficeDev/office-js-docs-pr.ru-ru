---
title: Элемент Method в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: fded84344182bb45597b00a794f18defaa44d3b3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432825"
---
# <a name="method-element"></a><span data-ttu-id="76bbc-102">Элемент Method</span><span class="sxs-lookup"><span data-stu-id="76bbc-102">Method element</span></span>

<span data-ttu-id="76bbc-103">Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="76bbc-103">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="76bbc-104">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="76bbc-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="76bbc-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="76bbc-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="76bbc-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="76bbc-106">Contained in</span></span>

[<span data-ttu-id="76bbc-107">Методы</span><span class="sxs-lookup"><span data-stu-id="76bbc-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="76bbc-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="76bbc-108">Attributes</span></span>

|<span data-ttu-id="76bbc-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="76bbc-109">**Attribute**</span></span>|<span data-ttu-id="76bbc-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="76bbc-110">**Type**</span></span>|<span data-ttu-id="76bbc-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="76bbc-111">**Required**</span></span>|<span data-ttu-id="76bbc-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="76bbc-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="76bbc-113">Имя</span><span class="sxs-lookup"><span data-stu-id="76bbc-113">Name</span></span>|<span data-ttu-id="76bbc-114">string</span><span class="sxs-lookup"><span data-stu-id="76bbc-114">string</span></span>|<span data-ttu-id="76bbc-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="76bbc-115">required</span></span>|<span data-ttu-id="76bbc-p101">Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы задать метод **getSelectedDataAsync**, необходимо указать `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="76bbc-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="76bbc-118">Примечания</span><span class="sxs-lookup"><span data-stu-id="76bbc-118">Remarks</span></span>

<span data-ttu-id="76bbc-119">Элементы **Methods** и **Method** не поддерживаются для почтовых надстроек. Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="76bbc-119">The  Methods and Method elements aren't supported by mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="76bbc-120">Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**.</span><span class="sxs-lookup"><span data-stu-id="76bbc-120">Important  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see Understanding the JavaScript API for Office.</span></span> <span data-ttu-id="76bbc-121">Дополнительные сведения о том, как это сделать, см. в статье [Общие сведения об API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="76bbc-121">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

