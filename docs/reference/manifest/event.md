# <a name="event-element"></a><span data-ttu-id="8409f-101">Элемент Event</span><span class="sxs-lookup"><span data-stu-id="8409f-101">Event element</span></span>

<span data-ttu-id="8409f-102">Определяет обработчик событий в надстройке.</span><span class="sxs-lookup"><span data-stu-id="8409f-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="8409f-103">Примечание. В настоящее время элемент `Event` поддерживается только в Outlook в Интернете из Office 365.</span><span class="sxs-lookup"><span data-stu-id="8409f-103">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="8409f-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8409f-104">Attributes</span></span>

|  <span data-ttu-id="8409f-105">Атрибут</span><span class="sxs-lookup"><span data-stu-id="8409f-105">Attribute</span></span>  |  <span data-ttu-id="8409f-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8409f-106">Required</span></span>  |  <span data-ttu-id="8409f-107">Описание</span><span class="sxs-lookup"><span data-stu-id="8409f-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8409f-108">Тип</span><span class="sxs-lookup"><span data-stu-id="8409f-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="8409f-109">Да</span><span class="sxs-lookup"><span data-stu-id="8409f-109">Yes</span></span>  | <span data-ttu-id="8409f-110">Задает обрабатываемое событие.</span><span class="sxs-lookup"><span data-stu-id="8409f-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="8409f-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="8409f-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="8409f-112">Да</span><span class="sxs-lookup"><span data-stu-id="8409f-112">Yes</span></span>  | <span data-ttu-id="8409f-p101">Задает способ выполнения обработчика событий (асинхронное или синхронное). В настоящее время поддерживаются только синхронные обработчики событий.</span><span class="sxs-lookup"><span data-stu-id="8409f-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="8409f-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="8409f-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="8409f-116">Да</span><span class="sxs-lookup"><span data-stu-id="8409f-116">Yes</span></span>  | <span data-ttu-id="8409f-117">Задает имя функции для обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="8409f-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="8409f-118">Атрибут Type</span><span class="sxs-lookup"><span data-stu-id="8409f-118">Type attribute</span></span>

<span data-ttu-id="8409f-p102">Обязательный. Указывает событие, при возникновении которого вызывается обработчик событий. В приведенной ниже таблице представлены допустимые значения этого атрибута.</span><span class="sxs-lookup"><span data-stu-id="8409f-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="8409f-122">Тип события</span><span class="sxs-lookup"><span data-stu-id="8409f-122">Event type</span></span>  |  <span data-ttu-id="8409f-123">Описание</span><span class="sxs-lookup"><span data-stu-id="8409f-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="8409f-124">Обработчик события будет вызван, когда пользователь отправляет сообщение или приглашение на собрание.</span><span class="sxs-lookup"><span data-stu-id="8409f-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="8409f-125">Атрибут FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="8409f-125">FunctionExecution attribute</span></span>

<span data-ttu-id="8409f-126">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="8409f-126">Required.</span></span> <span data-ttu-id="8409f-127">ОБЯЗАТЕЛЬНО указать значение `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="8409f-127">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="8409f-128">Атрибут FunctionName</span><span class="sxs-lookup"><span data-stu-id="8409f-128">FunctionName attribute</span></span>

<span data-ttu-id="8409f-p104">Обязательный. Задает имя функции для обработчика событий. Это значение должно совпадать с именем функции в [файле функции](functionfile.md)  надстройки.</span><span class="sxs-lookup"><span data-stu-id="8409f-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```