---
title: Элемент Event в файле манифеста
description: Определяет обработчик событий в надстройке.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 02037a54ad4b7e91a3697b53b04fa30e8a4909a9
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718232"
---
# <a name="event-element"></a><span data-ttu-id="02059-103">Элемент Event</span><span class="sxs-lookup"><span data-stu-id="02059-103">Event element</span></span>

<span data-ttu-id="02059-104">Определяет обработчик событий в надстройке.</span><span class="sxs-lookup"><span data-stu-id="02059-104">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="02059-105">В `Event` настоящее время элемент поддерживается только в Outlook в Интернете в Office 365.</span><span class="sxs-lookup"><span data-stu-id="02059-105">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="02059-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="02059-106">Attributes</span></span>

|  <span data-ttu-id="02059-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="02059-107">Attribute</span></span>  |  <span data-ttu-id="02059-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="02059-108">Required</span></span>  |  <span data-ttu-id="02059-109">Описание</span><span class="sxs-lookup"><span data-stu-id="02059-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="02059-110">Тип</span><span class="sxs-lookup"><span data-stu-id="02059-110">Type</span></span>](#type-attribute)  |  <span data-ttu-id="02059-111">Да</span><span class="sxs-lookup"><span data-stu-id="02059-111">Yes</span></span>  | <span data-ttu-id="02059-112">Задает обрабатываемое событие.</span><span class="sxs-lookup"><span data-stu-id="02059-112">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="02059-113">функтионексекутион</span><span class="sxs-lookup"><span data-stu-id="02059-113">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="02059-114">Да</span><span class="sxs-lookup"><span data-stu-id="02059-114">Yes</span></span>  | <span data-ttu-id="02059-p101">Задает способ выполнения обработчика событий (асинхронное или синхронное). В настоящее время поддерживаются только синхронные обработчики событий.</span><span class="sxs-lookup"><span data-stu-id="02059-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="02059-117">FunctionName</span><span class="sxs-lookup"><span data-stu-id="02059-117">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="02059-118">Да</span><span class="sxs-lookup"><span data-stu-id="02059-118">Yes</span></span>  | <span data-ttu-id="02059-119">Задает имя функции для обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="02059-119">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="02059-120">Атрибут Type</span><span class="sxs-lookup"><span data-stu-id="02059-120">Type attribute</span></span>

<span data-ttu-id="02059-p102">Обязательный. Указывает событие, при возникновении которого вызывается обработчик. В приведенной ниже таблице представлены допустимые значения этого атрибута.</span><span class="sxs-lookup"><span data-stu-id="02059-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="02059-124">Тип события</span><span class="sxs-lookup"><span data-stu-id="02059-124">Event type</span></span>  |  <span data-ttu-id="02059-125">Описание</span><span class="sxs-lookup"><span data-stu-id="02059-125">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="02059-126">Обработчик события будет вызван, когда пользователь отправляет сообщение или приглашение на собрание.</span><span class="sxs-lookup"><span data-stu-id="02059-126">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="02059-127">Атрибут FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="02059-127">FunctionExecution attribute</span></span>

<span data-ttu-id="02059-128">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="02059-128">Required.</span></span> <span data-ttu-id="02059-129">ДОЛЖНО быть задано значение `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="02059-129">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="02059-130">Атрибут FunctionName</span><span class="sxs-lookup"><span data-stu-id="02059-130">FunctionName attribute</span></span>

<span data-ttu-id="02059-p104">Обязательный. Задает имя функции для обработчика событий. Это значение должно совпадать с именем функции в [файле функции](functionfile.md) надстройки.</span><span class="sxs-lookup"><span data-stu-id="02059-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```
