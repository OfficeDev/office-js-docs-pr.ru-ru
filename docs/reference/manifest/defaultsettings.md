# <a name="defaultsettings-element"></a><span data-ttu-id="cf72b-101">Элемент DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="cf72b-101">DefaultSettings element</span></span>

<span data-ttu-id="cf72b-102">Указывает исходное расположение по умолчанию и другие параметры по умолчанию для контентной надстройки или надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="cf72b-102">Specifies the default source location and other default settings for your content or task pane add-in .</span></span>

<span data-ttu-id="cf72b-103">**Тип надстройки:** контентные надстройки и надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="cf72b-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="cf72b-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="cf72b-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="cf72b-105">Элемент, в котором содержится</span><span class="sxs-lookup"><span data-stu-id="cf72b-105">Contained in:</span></span>

[<span data-ttu-id="cf72b-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="cf72b-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="cf72b-107">Может содержать</span><span class="sxs-lookup"><span data-stu-id="cf72b-107">Can contain:</span></span>

|<span data-ttu-id="cf72b-108">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="cf72b-108">**Element**</span></span>|<span data-ttu-id="cf72b-109">**Контентные**</span><span class="sxs-lookup"><span data-stu-id="cf72b-109">**Content**</span></span>|<span data-ttu-id="cf72b-110">**Почтовые**</span><span class="sxs-lookup"><span data-stu-id="cf72b-110">**Mail**</span></span>|<span data-ttu-id="cf72b-111">**Области задач**</span><span class="sxs-lookup"><span data-stu-id="cf72b-111">\*\*\*\* Taskpane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="cf72b-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="cf72b-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="cf72b-113">x</span><span class="sxs-lookup"><span data-stu-id="cf72b-113">x</span></span>||<span data-ttu-id="cf72b-114">x</span><span class="sxs-lookup"><span data-stu-id="cf72b-114">x</span></span>|
|[<span data-ttu-id="cf72b-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="cf72b-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="cf72b-116">x</span><span class="sxs-lookup"><span data-stu-id="cf72b-116">x</span></span>|||
|[<span data-ttu-id="cf72b-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="cf72b-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="cf72b-118">x</span><span class="sxs-lookup"><span data-stu-id="cf72b-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="cf72b-119">Замечания</span><span class="sxs-lookup"><span data-stu-id="cf72b-119">Remarks</span></span>

<span data-ttu-id="cf72b-120">Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к надстройкам области задач и контентным надстройкам. В случае почтовых надстроек следует задавать расположения по умолчанию для исходных файлов и другие параметры по умолчанию с помощью элемента [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="cf72b-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

