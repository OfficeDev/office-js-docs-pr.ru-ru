# <a name="sets-element"></a><span data-ttu-id="d7278-101">Элемент Sets</span><span class="sxs-lookup"><span data-stu-id="d7278-101">Sets element</span></span>

<span data-ttu-id="d7278-102">Указывает минимальное подмножество API JavaScript для Office, необходимое для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="d7278-102">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="d7278-103">**Тип надстройки:** содержимое, область задач, почта</span><span class="sxs-lookup"><span data-stu-id="d7278-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d7278-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="d7278-104">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="d7278-105">Содержится в</span><span class="sxs-lookup"><span data-stu-id="d7278-105">Contained in:</span></span>

[<span data-ttu-id="d7278-106">Требования</span><span class="sxs-lookup"><span data-stu-id="d7278-106">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="d7278-107">Может содержать</span><span class="sxs-lookup"><span data-stu-id="d7278-107">Can contain:</span></span>

[<span data-ttu-id="d7278-108">Множество</span><span class="sxs-lookup"><span data-stu-id="d7278-108">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="d7278-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d7278-109">Attributes</span></span>

|<span data-ttu-id="d7278-110">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="d7278-110">**Attribute**</span></span>|<span data-ttu-id="d7278-111">**Тип**</span><span class="sxs-lookup"><span data-stu-id="d7278-111">**Type**</span></span>|<span data-ttu-id="d7278-112">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="d7278-112">**Required**</span></span>|<span data-ttu-id="d7278-113">**Описание**</span><span class="sxs-lookup"><span data-stu-id="d7278-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d7278-114">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="d7278-114">DefaultMinVersion</span></span>|<span data-ttu-id="d7278-115">string</span><span class="sxs-lookup"><span data-stu-id="d7278-115">string</span></span>|<span data-ttu-id="d7278-116">необязательный</span><span class="sxs-lookup"><span data-stu-id="d7278-116">optional</span></span>|<span data-ttu-id="d7278-p101">Задает значение атрибута **MinVersion** по умолчанию для всех дочерних элементов [Set](set.md). Значение по умолчанию: "1.1".</span><span class="sxs-lookup"><span data-stu-id="d7278-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="d7278-119">Замечания</span><span class="sxs-lookup"><span data-stu-id="d7278-119">Remarks</span></span>

<span data-ttu-id="d7278-120">Дополнительные сведения о наборах обязательных элементов см. в статье [Версии и наборы обязательных элементов Office](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d7278-120">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="d7278-121">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="d7278-121">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

