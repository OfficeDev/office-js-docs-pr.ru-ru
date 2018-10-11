# <a name="method-element"></a><span data-ttu-id="28c2a-101">Элемент Method</span><span class="sxs-lookup"><span data-stu-id="28c2a-101">Method element</span></span>

<span data-ttu-id="28c2a-102">Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="28c2a-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="28c2a-103">**Тип надстройки:** содержимое, область задач.</span><span class="sxs-lookup"><span data-stu-id="28c2a-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="28c2a-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="28c2a-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="28c2a-105">Содержится в</span><span class="sxs-lookup"><span data-stu-id="28c2a-105">Contained in:</span></span>

[<span data-ttu-id="28c2a-106">Методы</span><span class="sxs-lookup"><span data-stu-id="28c2a-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="28c2a-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="28c2a-107">Attributes</span></span>

|<span data-ttu-id="28c2a-108">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="28c2a-108">**Attribute**</span></span>|<span data-ttu-id="28c2a-109">**Тип**</span><span class="sxs-lookup"><span data-stu-id="28c2a-109">**Type**</span></span>|<span data-ttu-id="28c2a-110">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="28c2a-110">**Required**</span></span>|<span data-ttu-id="28c2a-111">**Описание**</span><span class="sxs-lookup"><span data-stu-id="28c2a-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="28c2a-112">Имя</span><span class="sxs-lookup"><span data-stu-id="28c2a-112">Name</span></span>|<span data-ttu-id="28c2a-113">string</span><span class="sxs-lookup"><span data-stu-id="28c2a-113">string</span></span>|<span data-ttu-id="28c2a-114">обязательный</span><span class="sxs-lookup"><span data-stu-id="28c2a-114">required</span></span>|<span data-ttu-id="28c2a-p101">Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы задать метод **getSelectedDataAsync**, необходимо указать `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="28c2a-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="28c2a-117">Замечания</span><span class="sxs-lookup"><span data-stu-id="28c2a-117">Remarks</span></span>

<span data-ttu-id="28c2a-118">Элементы  **Methods** и **Method** не поддерживаются надстройками почты. Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="28c2a-118">The  Methods and Method elements aren't supported by mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="28c2a-119">Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**.</span><span class="sxs-lookup"><span data-stu-id="28c2a-119">Important  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see Understanding the JavaScript API for Office.</span></span> <span data-ttu-id="28c2a-120">Дополнительные сведения о том, как это сделать, см. в статье [Общие сведения об API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="28c2a-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

