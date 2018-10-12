# <a name="set-element"></a><span data-ttu-id="8b2db-101">Элемент Set</span><span class="sxs-lookup"><span data-stu-id="8b2db-101">Set element</span></span>

<span data-ttu-id="8b2db-102">Указывает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="8b2db-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="8b2db-103">**Тип надстройки:** содержимое, область задач, почта</span><span class="sxs-lookup"><span data-stu-id="8b2db-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8b2db-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="8b2db-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="8b2db-105">Содержится в</span><span class="sxs-lookup"><span data-stu-id="8b2db-105">Contained in:</span></span>

[<span data-ttu-id="8b2db-106">Наборы</span><span class="sxs-lookup"><span data-stu-id="8b2db-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="8b2db-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b2db-107">Attributes</span></span>

|<span data-ttu-id="8b2db-108">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="8b2db-108">**Attribute**</span></span>|<span data-ttu-id="8b2db-109">**Тип**</span><span class="sxs-lookup"><span data-stu-id="8b2db-109">**Type**</span></span>|<span data-ttu-id="8b2db-110">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="8b2db-110">**Required**</span></span>|<span data-ttu-id="8b2db-111">**Описание**</span><span class="sxs-lookup"><span data-stu-id="8b2db-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="8b2db-112">Имя</span><span class="sxs-lookup"><span data-stu-id="8b2db-112">Name</span></span>|<span data-ttu-id="8b2db-113">string</span><span class="sxs-lookup"><span data-stu-id="8b2db-113">string</span></span>|<span data-ttu-id="8b2db-114">обязательный</span><span class="sxs-lookup"><span data-stu-id="8b2db-114">required</span></span>|<span data-ttu-id="8b2db-115">Имя [набора требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="8b2db-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="8b2db-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="8b2db-116">MinVersion</span></span>|<span data-ttu-id="8b2db-117">string</span><span class="sxs-lookup"><span data-stu-id="8b2db-117">string</span></span>|<span data-ttu-id="8b2db-118">необязательный</span><span class="sxs-lookup"><span data-stu-id="8b2db-118">optional</span></span>|<span data-ttu-id="8b2db-p101">Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion**, если оно указано в родительском элементе [Sets](sets.md).</span><span class="sxs-lookup"><span data-stu-id="8b2db-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="8b2db-121">Замечания</span><span class="sxs-lookup"><span data-stu-id="8b2db-121">Remarks</span></span>

<span data-ttu-id="8b2db-122">Дополнительные сведения о наборах требований см. в статье [версии и наборы требований  Office](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="8b2db-122">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="8b2db-123">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="8b2db-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="8b2db-124">Для надстроек почты существует только один `"Mailbox"` набор требований.</span><span class="sxs-lookup"><span data-stu-id="8b2db-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="8b2db-125">Этот набор обязательных элементов содержит целое подмножество API, поддерживаемых в надстройках почты для Outlook, и необходимо указать `"Mailbox"` набор обязательных требований в манифесте надстройки почты (это необходимо, как в случае содержимого, так и надстройки области задач).</span><span class="sxs-lookup"><span data-stu-id="8b2db-125">Important  For mail add-ins, there is only one   requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.</span></span> <span data-ttu-id="8b2db-126">Кроме того, невозможно объявить поддержку для отдельных методов в надстройках почты.</span><span class="sxs-lookup"><span data-stu-id="8b2db-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
