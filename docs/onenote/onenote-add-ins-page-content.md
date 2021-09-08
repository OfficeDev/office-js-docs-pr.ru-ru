---
title: Работа с содержимым страницы в OneNote
description: Узнайте, как работать с OneNote контентом страницы с помощью API JavaScript.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 01aa4a65f6f1d7ae8fccf490986c10035d30b0c3
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938584"
---
# <a name="work-with-onenote-page-content"></a>Работа с содержимым страницы в OneNote

В API JavaScript для надстроек OneNote содержимое страницы представлено указанной ниже объектной моделью.

  ![OneNote схема объектной модели страницы.](../images/one-note-om-page.png)

- Объект Page содержит коллекцию объектов PageContent.
- Объект PageContent содержит контент типов Outline, Image или Other.
- Объект Outline содержит коллекцию объектов Paragraph.
- Объект Paragraph содержит контент типов RichText, Image, Table или Other.

Чтобы создать пустую OneNote страницу, используйте один из следующих методов.

- [Section.addPage](/javascript/api/onenote/onenote.section#addPage_title_)
- [Page.insertPageAsSibling](/javascript/api/onenote/onenote.section#insertSectionAsSibling_location__title_)

Затем используйте методы в указанных ниже объектах для работы с содержимым страницы, например `Page.addOutline` и `Outline.appendHtml`.

- [Страница](/javascript/api/onenote/onenote.page)
- [Outline](/javascript/api/onenote/onenote.outline)
- [Paragraph](/javascript/api/onenote/onenote.paragraph)

Для представления содержимого и структуры страницы OneNote используется HTML. Для создания или обновления содержимого страницы поддерживается только подмножество HTML, как описано ниже.

## <a name="supported-html"></a>Поддерживаемые элементы HTML

API OneNote JavaScript поддерживает следующий HTML для создания и обновления контента страницы.

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>`
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>`
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

> [!NOTE]
> Импорт HTML в OneNote консолидирует пробелы. Полученное в результате содержимое вставляется в одну структуру.

Приложение OneNote пытается наилучшим образом преобразовать HTML в содержимое страницы, обеспечивая безопасность для пользователей. Так как стандарты HTML и CSS не полностью соответствуют модели содержимого OneNote, будут иметься различия во внешнем виде, особенно при использовании стилей CSS. Рекомендуется использовать объекты JavaScript, если требуется определенное форматирование.

## <a name="accessing-page-contents"></a>Доступ к содержимому страницы

Через  доступ можно получить только к `Page#load`. Чтобы изменить активную страницу, вызовите команду `navigateToPage($page)`.

Метаданные, например "Название", можно запросить для любой страницы.

## <a name="see-also"></a>См. также

- [Обзор API JavaScript для OneNote](onenote-add-ins-programming-overview.md)
- [Справочник по API JavaScript для OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
