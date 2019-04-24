---
title: Элемент Method в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 19234b35e1faf8a8cc52a9e893fcc720793cadae
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450655"
---
# <a name="method-element"></a><span data-ttu-id="0ec1b-102">Элемент Method</span><span class="sxs-lookup"><span data-stu-id="0ec1b-102">Method element</span></span>

<span data-ttu-id="0ec1b-103">Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="0ec1b-103">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="0ec1b-104">**Тип надстройки:** контентные надстройки и надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="0ec1b-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="0ec1b-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="0ec1b-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="0ec1b-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="0ec1b-106">Contained in</span></span>

[<span data-ttu-id="0ec1b-107">Методы</span><span class="sxs-lookup"><span data-stu-id="0ec1b-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="0ec1b-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0ec1b-108">Attributes</span></span>

|<span data-ttu-id="0ec1b-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="0ec1b-109">**Attribute**</span></span>|<span data-ttu-id="0ec1b-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="0ec1b-110">**Type**</span></span>|<span data-ttu-id="0ec1b-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="0ec1b-111">**Required**</span></span>|<span data-ttu-id="0ec1b-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0ec1b-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0ec1b-113">Имя</span><span class="sxs-lookup"><span data-stu-id="0ec1b-113">Name</span></span>|<span data-ttu-id="0ec1b-114">string</span><span class="sxs-lookup"><span data-stu-id="0ec1b-114">string</span></span>|<span data-ttu-id="0ec1b-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0ec1b-115">required</span></span>|<span data-ttu-id="0ec1b-p101">Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы задать метод **getSelectedDataAsync**, необходимо указать `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="0ec1b-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="0ec1b-118">Примечания</span><span class="sxs-lookup"><span data-stu-id="0ec1b-118">Remarks</span></span>

<span data-ttu-id="0ec1b-119">Элементы **Methods** и **Method** не поддерживаются для почтовых надстроек. Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="0ec1b-119">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="0ec1b-120">Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**.</span><span class="sxs-lookup"><span data-stu-id="0ec1b-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="0ec1b-121">Дополнительные сведения о том, как это сделать, см. в статье [Общие сведения об API JavaScript для Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="0ec1b-121">For more information about how to do this, see [Understanding the JavaScript API for Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

