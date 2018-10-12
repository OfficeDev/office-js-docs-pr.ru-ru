# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="67117-101">Наборы требований API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="67117-101">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="67117-102">Наборы требований — это именованные группы требований API.</span><span class="sxs-lookup"><span data-stu-id="67117-102">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="67117-103">С помощью наборов требований, указанных в манифесте, или проверки в среде выполнения надстройки Office определяют, поддерживает ли ведущее приложение Office необходимые API.</span><span class="sxs-lookup"><span data-stu-id="67117-103">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see Specify Office hosts and API requirements.</span></span> <span data-ttu-id="67117-104">Дополнительные сведения см. в статье [Версии Office и наборы требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="67117-104">For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="67117-105">В приведенной ниже таблице перечислены наборы требований для OneNote, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.</span><span class="sxs-lookup"><span data-stu-id="67117-105">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="67117-106">Набор требований</span><span class="sxs-lookup"><span data-stu-id="67117-106">Requirement set</span></span>  |  <span data-ttu-id="67117-107">Office Online</span><span class="sxs-lookup"><span data-stu-id="67117-107">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="67117-108">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="67117-108">OneNoteApi 1.1</span></span>  | <span data-ttu-id="67117-109">Сентябрь 2016 г.</span><span class="sxs-lookup"><span data-stu-id="67117-109">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="67117-110">Стандартные наборы требований API для Office</span><span class="sxs-lookup"><span data-stu-id="67117-110">Office common API requirement sets</span></span>

<span data-ttu-id="67117-111">Сведения о стандартных наборах требований API см. в статье [Стандартные наборы требований API для Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="67117-111">For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="67117-112">API JavaScript для OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="67117-112">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="67117-113">API JavaScript для OneNote 1.1 — первая версия этого API.</span><span class="sxs-lookup"><span data-stu-id="67117-113">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="67117-114">Подробнее об API см. [Общие сведения о программировании API JavaScript для OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="67117-114">For details about the API, see the [OneNote JavaScript API](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview) reference topics.</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="67117-115">Проверка поддержки требований в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="67117-115">Runtime requirement support check</span></span>

<span data-ttu-id="67117-116">Во время выполнения кода надстройки могут проверять, поддерживает ли ведущее приложение набор требований API, выполняя следующую проверку:</span><span class="sxs-lookup"><span data-stu-id="67117-116">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="67117-117">Проверка поддержки требований в манифесте</span><span class="sxs-lookup"><span data-stu-id="67117-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="67117-p103">Используйте элемент Requirements в манифесте надстройки, чтобы указать ключевые наборы требований или элементы API, которые должна использовать надстройка. Если платформа или ведущее приложение Office не поддерживает наборы требований или элементы API, указанные в элементе Requirements, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе «Мои надстройки».</span><span class="sxs-lookup"><span data-stu-id="67117-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="67117-120">Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор требований OneNoteApi версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="67117-120">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="67117-121">См. также</span><span class="sxs-lookup"><span data-stu-id="67117-121">See also</span></span>

- [<span data-ttu-id="67117-122">Версии Office и наборы требований</span><span class="sxs-lookup"><span data-stu-id="67117-122">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="67117-123">Указание ведущих приложений Office и требований API</span><span class="sxs-lookup"><span data-stu-id="67117-123">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="67117-124">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="67117-124">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
