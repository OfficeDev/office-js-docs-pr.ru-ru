---
title: Элемент Event в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: eda895b01e106d67eef70f199be64086e9372bef
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432741"
---
# <a name="event-element"></a><span data-ttu-id="19ecc-102">Элемент Event</span><span class="sxs-lookup"><span data-stu-id="19ecc-102">Event element</span></span>

<span data-ttu-id="19ecc-103">Определяет обработчик событий в надстройке.</span><span class="sxs-lookup"><span data-stu-id="19ecc-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="19ecc-104">В настоящее время элемент `Event` поддерживается только в Outlook в Интернете из Office 365.</span><span class="sxs-lookup"><span data-stu-id="19ecc-104">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="19ecc-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19ecc-105">Attributes</span></span>

|  <span data-ttu-id="19ecc-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="19ecc-106">Attribute</span></span>  |  <span data-ttu-id="19ecc-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="19ecc-107">Required</span></span>  |  <span data-ttu-id="19ecc-108">Описание</span><span class="sxs-lookup"><span data-stu-id="19ecc-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="19ecc-109">Type</span><span class="sxs-lookup"><span data-stu-id="19ecc-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="19ecc-110">Да</span><span class="sxs-lookup"><span data-stu-id="19ecc-110">Yes</span></span>  | <span data-ttu-id="19ecc-111">Задает обрабатываемое событие.</span><span class="sxs-lookup"><span data-stu-id="19ecc-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="19ecc-112">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="19ecc-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="19ecc-113">ДА</span><span class="sxs-lookup"><span data-stu-id="19ecc-113">Yes</span></span>  | <span data-ttu-id="19ecc-p101">Задает способ выполнения обработчика событий (асинхронное или синхронное). В настоящее время поддерживаются только синхронные обработчики событий.</span><span class="sxs-lookup"><span data-stu-id="19ecc-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="19ecc-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="19ecc-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="19ecc-117">Да</span><span class="sxs-lookup"><span data-stu-id="19ecc-117">Yes</span></span>  | <span data-ttu-id="19ecc-118">Задает имя функции для обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="19ecc-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="19ecc-119">Атрибут Type</span><span class="sxs-lookup"><span data-stu-id="19ecc-119">Type attribute</span></span>

<span data-ttu-id="19ecc-p102">Обязательный. Указывает событие, при возникновении которого вызывается обработчик. В приведенной ниже таблице представлены допустимые значения этого атрибута.</span><span class="sxs-lookup"><span data-stu-id="19ecc-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="19ecc-123">Тип события</span><span class="sxs-lookup"><span data-stu-id="19ecc-123">Event type</span></span>  |  <span data-ttu-id="19ecc-124">Описание</span><span class="sxs-lookup"><span data-stu-id="19ecc-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="19ecc-125">Обработчик события будет вызван, когда пользователь отправляет сообщение или приглашение на собрание.</span><span class="sxs-lookup"><span data-stu-id="19ecc-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="19ecc-126">Атрибут FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="19ecc-126">FunctionExecution attribute</span></span>

<span data-ttu-id="19ecc-p103">Обязательный. ДОЛЖНО быть задано значение `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="19ecc-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="19ecc-129">Атрибут FunctionName</span><span class="sxs-lookup"><span data-stu-id="19ecc-129">FunctionName attribute</span></span>

<span data-ttu-id="19ecc-p104">Обязательный. Задает имя функции для обработчика событий. Это значение должно совпадать с именем функции в [файле функции](functionfile.md) надстройки.</span><span class="sxs-lookup"><span data-stu-id="19ecc-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```