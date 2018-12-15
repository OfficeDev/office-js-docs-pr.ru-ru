# <a name="appdomain-element"></a><span data-ttu-id="63328-101">Элемент AppDomain</span><span class="sxs-lookup"><span data-stu-id="63328-101">AppDomain element</span></span>

<span data-ttu-id="63328-102">Указывает дополнительный домен, который будет использоваться для загрузки страниц в окне надстройки.</span><span class="sxs-lookup"><span data-stu-id="63328-102">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="63328-103">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="63328-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="63328-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="63328-104">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="63328-105">Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="63328-105">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="63328-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="63328-106">Contained in</span></span>

[<span data-ttu-id="63328-107">AppDomains</span><span class="sxs-lookup"><span data-stu-id="63328-107">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="63328-108">Примечания</span><span class="sxs-lookup"><span data-stu-id="63328-108">Remarks</span></span>

<span data-ttu-id="63328-109">Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="63328-109">The  AppDomains and **AppDomain** elements are used to specify any additional domains other than the one specified in the [SourceLocation element. For more information, see Office Add-ins XML manifest](sourcelocation.md).</span></span> <span data-ttu-id="63328-110">Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="63328-110">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
