---
title: Элемент Event в файле манифеста
description: Определяет обработчик событий в надстройке.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 3d8e94c10bed214dd976b3048e11328f10f99325
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611549"
---
# <a name="event-element"></a><span data-ttu-id="4a82b-103">Элемент Event</span><span class="sxs-lookup"><span data-stu-id="4a82b-103">Event element</span></span>

<span data-ttu-id="4a82b-104">Определяет обработчик событий в надстройке.</span><span class="sxs-lookup"><span data-stu-id="4a82b-104">Defines an event handler in an add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="4a82b-105">Сведения о поддержке и использовании можно найти [в статье функция отправки почты для надстроек Outlook](../../outlook/outlook-on-send-addins.md).</span><span class="sxs-lookup"><span data-stu-id="4a82b-105">For information about support and usage, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="4a82b-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4a82b-106">Attributes</span></span>

|  <span data-ttu-id="4a82b-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="4a82b-107">Attribute</span></span>  |  <span data-ttu-id="4a82b-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4a82b-108">Required</span></span>  |  <span data-ttu-id="4a82b-109">Описание</span><span class="sxs-lookup"><span data-stu-id="4a82b-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4a82b-110">Тип</span><span class="sxs-lookup"><span data-stu-id="4a82b-110">Type</span></span>](#type-attribute)  |  <span data-ttu-id="4a82b-111">Да</span><span class="sxs-lookup"><span data-stu-id="4a82b-111">Yes</span></span>  | <span data-ttu-id="4a82b-112">Задает обрабатываемое событие.</span><span class="sxs-lookup"><span data-stu-id="4a82b-112">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="4a82b-113">функтионексекутион</span><span class="sxs-lookup"><span data-stu-id="4a82b-113">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="4a82b-114">Да</span><span class="sxs-lookup"><span data-stu-id="4a82b-114">Yes</span></span>  | <span data-ttu-id="4a82b-p101">Задает способ выполнения обработчика событий (асинхронное или синхронное). В настоящее время поддерживаются только синхронные обработчики событий.</span><span class="sxs-lookup"><span data-stu-id="4a82b-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="4a82b-117">FunctionName</span><span class="sxs-lookup"><span data-stu-id="4a82b-117">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="4a82b-118">Да</span><span class="sxs-lookup"><span data-stu-id="4a82b-118">Yes</span></span>  | <span data-ttu-id="4a82b-119">Задает имя функции для обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="4a82b-119">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="4a82b-120">Атрибут Type</span><span class="sxs-lookup"><span data-stu-id="4a82b-120">Type attribute</span></span>

<span data-ttu-id="4a82b-p102">Обязательный. Указывает событие, при возникновении которого вызывается обработчик. В приведенной ниже таблице представлены допустимые значения этого атрибута.</span><span class="sxs-lookup"><span data-stu-id="4a82b-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="4a82b-124">Тип события</span><span class="sxs-lookup"><span data-stu-id="4a82b-124">Event type</span></span>  |  <span data-ttu-id="4a82b-125">Описание</span><span class="sxs-lookup"><span data-stu-id="4a82b-125">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="4a82b-126">Обработчик события будет вызван, когда пользователь отправляет сообщение или приглашение на собрание.</span><span class="sxs-lookup"><span data-stu-id="4a82b-126">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="4a82b-127">Атрибут FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="4a82b-127">FunctionExecution attribute</span></span>

<span data-ttu-id="4a82b-128">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="4a82b-128">Required.</span></span> <span data-ttu-id="4a82b-129">ДОЛЖНО быть задано значение `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="4a82b-129">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="4a82b-130">Атрибут FunctionName</span><span class="sxs-lookup"><span data-stu-id="4a82b-130">FunctionName attribute</span></span>

<span data-ttu-id="4a82b-p104">Обязательный. Задает имя функции для обработчика событий. Это значение должно совпадать с именем функции в [файле функции](functionfile.md) надстройки.</span><span class="sxs-lookup"><span data-stu-id="4a82b-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
