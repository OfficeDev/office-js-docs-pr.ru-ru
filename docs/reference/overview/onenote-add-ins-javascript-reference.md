---
title: Обзор API JavaScript для OneNote
description: ''
ms.date: 06/20/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 68ac6f94921ba3b1ea14f364988b57ef86809890
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127130"
---
# <a name="onenote-javascript-api-overview"></a>Обзор API JavaScript для OneNote

Область применения: OneNote в Интернете

Ниже приведены ссылки на высокоуровневые объекты OneNote, доступные в API. Каждая ссылка на страницу объекта содержит описание свойств, событий и методов, доступных для объекта. Чтобы узнать больше, перейдите по указанным ниже ссылкам. 
    
- [Application](/javascript/api/onenote/onenote.application): объект верхнего уровня, используемый для доступа ко всем глобально адресуемым объектам OneNote, таким как активная записная книжка и активный раздел.

- [Notebook](/javascript/api/onenote/onenote.notebook): записная книжка. Записные книжки содержат группы разделов и разделы.
    - [NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): представляет коллекцию записных книжек.

- [SectionGroup](/javascript/api/onenote/onenote.sectiongroup): группа разделов. Группы разделов содержат разделы и группы разделов.
    - [SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): коллекция групп разделов.

- [Section](/javascript/api/onenote/onenote.section): раздел. Разделы содержат страницы.
    - [SectionCollection](/javascript/api/onenote/onenote.sectioncollection): коллекция разделов.

- [Page](/javascript/api/onenote/onenote.page): страница. Страницы содержат объекты PageContent.
    - [PageCollection](/javascript/api/onenote/onenote.pagecollection): коллекция страниц.

- [PageContent](/javascript/api/onenote/onenote.pagecontent): область верхнего уровня на странице, содержащая контент, например типов Outline или Image. Объекту PageContent можно назначить позицию на странице.
    - [PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): коллекция объектов PageContent, представляющая содержимое страницы.

- [Outline](/javascript/api/onenote/onenote.outline): контейнер для объектов Paragraph. Объект Outline — прямой потомок объекта PageContent.

- [Image](/javascript/api/onenote/onenote.image): объект Image. Объект Image может быть прямым потомком объекта PageContent или объекта Paragraph.

- [Paragraph](/javascript/api/onenote/onenote.paragraph): Контейнер для содержимого, отображаемого на странице. Объект Paragraph — прямой потомок объекта Outline.
    - [ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): коллекция объектов Paragraph в объекте Outline.

- [RichText](/javascript/api/onenote/onenote.richtext): объект RichText.

- [Table](/javascript/api/onenote/onenote.table): контейнер для объектов TableRow.

- [TableRow](/javascript/api/onenote/onenote.tablerow): контейнер для объектов TableCell.
    - [TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): Коллекция объектов TableRow в объекте Table.
 
- [TableCell](/javascript/api/onenote/onenote.tablecell): контейнер для объектов Paragraph.
    - [TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection) коллекция объектов TableCell в объекте TableRow.

## <a name="onenote-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для OneNote

Наборы требований — это именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения о наборах обязательных элементов API JavaScript для OneNote см. в статье [Наборы обязательных элементов API JavaScript для OneNote](../requirement-sets/onenote-api-requirement-sets.md).

## <a name="onenote-javascript-api-reference"></a>Справочник по API JavaScript для OneNote

Дополнительные сведения об API JavaScript для OneNote см. в [справочной документации по API JavaScript для OneNote](/javascript/api/onenote).

## <a name="see-also"></a>См. также

- [Обзор API JavaScript для OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [Создание первой надстройки OneNote](../../quickstarts/onenote-quickstart.md)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](/office/dev/add-ins/overview/office-add-ins)
