# <a name="appdomains-element"></a><span data-ttu-id="b658f-101">Элемент AppDomains</span><span class="sxs-lookup"><span data-stu-id="b658f-101">AppDomains element</span></span>

<span data-ttu-id="b658f-p101">Определяет все домены, кроме указанного в элементе SourceLocation, которые надстройка Office будет использовать для загрузки страниц. Для каждого дополнительного домена укажите элемент AppDomain.</span><span class="sxs-lookup"><span data-stu-id="b658f-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="b658f-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="b658f-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b658f-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b658f-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="b658f-106">Значение каждого элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="b658f-106">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="b658f-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="b658f-107">Contained in</span></span>

[<span data-ttu-id="b658f-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b658f-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="b658f-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="b658f-109">Can contain</span></span>

[<span data-ttu-id="b658f-110">AppDomain</span><span class="sxs-lookup"><span data-stu-id="b658f-110">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="b658f-111">Примечания</span><span class="sxs-lookup"><span data-stu-id="b658f-111">Remarks</span></span>

<span data-ttu-id="b658f-112">По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="b658f-112">By default, your add-in can load any page that is in the same domain as the location specified in the SourceLocation element. To load pages that are not in the same domain as the add-in, specify the domains by using the AppDomains and AppDomain elements. This element can't be empty.</span></span> <span data-ttu-id="b658f-113">Для загрузки страниц из других доменов, укажите их домены в элементах **AppDomains** и **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="b658f-113">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="b658f-114">Этот элемент не может быть пустым.</span><span class="sxs-lookup"><span data-stu-id="b658f-114">This element can't be empty.</span></span>
