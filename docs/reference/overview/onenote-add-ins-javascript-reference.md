---
title: Обзор API JavaScript для OneNote
description: ''
ms.date: 10/09/2018
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: f8fed0104412f60ec59146ef7820be958047d1f3
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052744"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="9b36d-102">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="9b36d-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="9b36d-103">Область применения: OneNote Online</span><span class="sxs-lookup"><span data-stu-id="9b36d-103">Applies to: OneNote Online</span></span>

<span data-ttu-id="9b36d-104">Ниже приведены ссылки на высокоуровневые объекты OneNote, доступные в API.</span><span class="sxs-lookup"><span data-stu-id="9b36d-104">The following links show the high level OneNote objects available in the API.</span></span> <span data-ttu-id="9b36d-105">Каждая ссылка на страницу объекта содержит описание свойств, событий и методов, доступных для объекта.</span><span class="sxs-lookup"><span data-stu-id="9b36d-105">Each object page link contains a description of the properties, events, and methods available on the object.</span></span> <span data-ttu-id="9b36d-106">Чтобы узнать больше, перейдите по указанным ниже ссылкам.</span><span class="sxs-lookup"><span data-stu-id="9b36d-106">Explore these links to learn more.</span></span> 
    
- <span data-ttu-id="9b36d-107">[Application](/javascript/api/onenote/onenote.application): объект верхнего уровня, используемый для доступа ко всем глобально адресуемым объектам OneNote, таким как активная записная книжка и активный раздел.</span><span class="sxs-lookup"><span data-stu-id="9b36d-107">[Application](/javascript/api/onenote/onenote.application): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.</span></span>

- <span data-ttu-id="9b36d-p102">[Notebook](/javascript/api/onenote/onenote.notebook): записная книжка. Записные книжки содержат группы разделов и разделы.</span><span class="sxs-lookup"><span data-stu-id="9b36d-p102">[Notebook](/javascript/api/onenote/onenote.notebook): A notebook. Notebooks contain section groups and sections.</span></span>
    - <span data-ttu-id="9b36d-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): представляет коллекцию записных книжек.</span><span class="sxs-lookup"><span data-stu-id="9b36d-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): A collection of notebooks.</span></span>

- <span data-ttu-id="9b36d-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): группа разделов. Группы разделов содержат разделы и группы разделов.</span><span class="sxs-lookup"><span data-stu-id="9b36d-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): A section group. Section groups contain section groups and sections.</span></span>
    - <span data-ttu-id="9b36d-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): коллекция групп разделов.</span><span class="sxs-lookup"><span data-stu-id="9b36d-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): A collection of section groups.</span></span>

- <span data-ttu-id="9b36d-p104">[Section](/javascript/api/onenote/onenote.section): раздел. Разделы содержат страницы.</span><span class="sxs-lookup"><span data-stu-id="9b36d-p104">[Section](/javascript/api/onenote/onenote.section): A section. Sections contain pages.</span></span>
    - <span data-ttu-id="9b36d-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): коллекция разделов.</span><span class="sxs-lookup"><span data-stu-id="9b36d-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): A collection of sections.</span></span>

- <span data-ttu-id="9b36d-p105">[Page](/javascript/api/onenote/onenote.page): страница. Страницы содержат объекты PageContent.</span><span class="sxs-lookup"><span data-stu-id="9b36d-p105">[Page](/javascript/api/onenote/onenote.page): A page. Pages contain PageContent objects.</span></span>
    - <span data-ttu-id="9b36d-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection): коллекция страниц.</span><span class="sxs-lookup"><span data-stu-id="9b36d-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection): A collection of pages.</span></span>

- <span data-ttu-id="9b36d-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): область верхнего уровня на странице, содержащая контент, например типов Outline или Image. Объекту PageContent можно назначить позицию на странице.</span><span class="sxs-lookup"><span data-stu-id="9b36d-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.</span></span>
    - <span data-ttu-id="9b36d-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): коллекция объектов PageContent, представляющая содержимое страницы.</span><span class="sxs-lookup"><span data-stu-id="9b36d-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): A collection of PageContent objects, which represents the contents of a page.</span></span>

- <span data-ttu-id="9b36d-p107">[Outline](/javascript/api/onenote/onenote.outline): контейнер для объектов Paragraph. Объект Outline — прямой потомок объекта PageContent.</span><span class="sxs-lookup"><span data-stu-id="9b36d-p107">[Outline](/javascript/api/onenote/onenote.outline): A container for Paragraph objects. An Outline is a direct child of a PageContent object.</span></span>

- <span data-ttu-id="9b36d-p108">[Image](/javascript/api/onenote/onenote.image): объект Image. Объект Image может быть прямым потомком объекта PageContent или объекта Paragraph.</span><span class="sxs-lookup"><span data-stu-id="9b36d-p108">[Image](/javascript/api/onenote/onenote.image): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.</span></span>

- <span data-ttu-id="9b36d-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): контейнер для содержимого, отображаемого на странице. Объект Paragraph — прямой потомок объекта Outline.</span><span class="sxs-lookup"><span data-stu-id="9b36d-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): A container for the visible content on a page. A Paragraph is a direct child of an Outline.</span></span>
    - <span data-ttu-id="9b36d-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): коллекция объектов Paragraph в объекте Outline.</span><span class="sxs-lookup"><span data-stu-id="9b36d-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): A collection of Paragraph objects in an Outline.</span></span>

- <span data-ttu-id="9b36d-130">[RichText](/javascript/api/onenote/onenote.richtext): объект RichText.</span><span class="sxs-lookup"><span data-stu-id="9b36d-130">[RichText](/javascript/api/onenote/onenote.richtext): A RichText object.</span></span>

- <span data-ttu-id="9b36d-131">[Table](/javascript/api/onenote/onenote.table): контейнер для объектов TableRow.</span><span class="sxs-lookup"><span data-stu-id="9b36d-131">[Table](/javascript/api/onenote/onenote.table): A container for TableRow objects.</span></span>

- <span data-ttu-id="9b36d-132">[TableRow](/javascript/api/onenote/onenote.tablerow): контейнер для объектов TableCell.</span><span class="sxs-lookup"><span data-stu-id="9b36d-132">[TableRow](/javascript/api/onenote/onenote.tablerow): A container for TableCell objects.</span></span>
    - <span data-ttu-id="9b36d-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): Коллекция объектов TableRow в объекте Table.</span><span class="sxs-lookup"><span data-stu-id="9b36d-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): A collection of TableRow objects in a Table.</span></span>
 
- <span data-ttu-id="9b36d-134">[TableCell](/javascript/api/onenote/onenote.tablecell): контейнер для объектов Paragraph.</span><span class="sxs-lookup"><span data-stu-id="9b36d-134">[TableCell](/javascript/api/onenote/onenote.tablecell): A container for Paragraph objects.</span></span>
    - <span data-ttu-id="9b36d-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection) коллекция объектов TableCell в объекте TableRow.</span><span class="sxs-lookup"><span data-stu-id="9b36d-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): A collection of TableCell objects in a TableRow.</span></span>

## <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="9b36d-136">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="9b36d-136">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="9b36d-137">Наборы обязательных элементов — именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="9b36d-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="9b36d-138">Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API.</span><span class="sxs-lookup"><span data-stu-id="9b36d-138">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="9b36d-139">Дополнительные сведения о наборах обязательных элементов API JavaScript для OneNote см. в статье [Наборы обязательных элементов API JavaScript для OneNote](../requirement-sets/onenote-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="9b36d-139">For detailed information about OneNote JavaScript API requirement sets, see the [OneNote JavaScript API requirement sets](../requirement-sets/onenote-api-requirement-sets.md) article.</span></span>

## <a name="onenote-javascript-api-reference"></a><span data-ttu-id="9b36d-140">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="9b36d-140">OneNote JavaScript API reference</span></span>

<span data-ttu-id="9b36d-141">Дополнительные сведения об API JavaScript для OneNote см. в [справочной документации по API JavaScript для OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="9b36d-141">For detailed information about the OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="9b36d-142">См. также</span><span class="sxs-lookup"><span data-stu-id="9b36d-142">See also</span></span>

- [<span data-ttu-id="9b36d-143">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="9b36d-143">OneNote JavaScript API programming overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [<span data-ttu-id="9b36d-144">Создание первой надстройки OneNote</span><span class="sxs-lookup"><span data-stu-id="9b36d-144">Build your first OneNote add-in</span></span>](../../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="9b36d-145">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="9b36d-145">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="9b36d-146">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9b36d-146">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
