# <a name="requirements-element"></a><span data-ttu-id="f0966-101">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="f0966-101">Requirements element</span></span>

<span data-ttu-id="f0966-102">Указывает минимальный набор требований API JavaScript для Office ([набор требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), необходимых для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="f0966-102">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="f0966-103">**Тип надстройки:** содержимое, область задач, почта</span><span class="sxs-lookup"><span data-stu-id="f0966-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f0966-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="f0966-104">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="f0966-105">Элемент, в котором содержится</span><span class="sxs-lookup"><span data-stu-id="f0966-105">Contained in:</span></span>

[<span data-ttu-id="f0966-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f0966-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f0966-107">Может содержать</span><span class="sxs-lookup"><span data-stu-id="f0966-107">Can contain:</span></span>

|<span data-ttu-id="f0966-108">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="f0966-108">**Element**</span></span>|<span data-ttu-id="f0966-109">**Контентные**</span><span class="sxs-lookup"><span data-stu-id="f0966-109">**Content**</span></span>|<span data-ttu-id="f0966-110">**Почтовые**</span><span class="sxs-lookup"><span data-stu-id="f0966-110">**Mail**</span></span>|<span data-ttu-id="f0966-111">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="f0966-111">\*\*\*\* Taskpane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="f0966-112">Наборы</span><span class="sxs-lookup"><span data-stu-id="f0966-112">Sets</span></span>](sets.md)|<span data-ttu-id="f0966-113">x</span><span class="sxs-lookup"><span data-stu-id="f0966-113">x</span></span>|<span data-ttu-id="f0966-114">x</span><span class="sxs-lookup"><span data-stu-id="f0966-114">x</span></span>|<span data-ttu-id="f0966-115">x</span><span class="sxs-lookup"><span data-stu-id="f0966-115">x</span></span>|
|[<span data-ttu-id="f0966-116">Методы</span><span class="sxs-lookup"><span data-stu-id="f0966-116">Methods</span></span>](methods.md)|<span data-ttu-id="f0966-117">x</span><span class="sxs-lookup"><span data-stu-id="f0966-117">x</span></span>||<span data-ttu-id="f0966-118">x</span><span class="sxs-lookup"><span data-stu-id="f0966-118">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="f0966-119">Замечания</span><span class="sxs-lookup"><span data-stu-id="f0966-119">Remarks</span></span>

<span data-ttu-id="f0966-120">Дополнительные сведения о наборах требований см. в статье [версии и наборы требований Office](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="f0966-120">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

