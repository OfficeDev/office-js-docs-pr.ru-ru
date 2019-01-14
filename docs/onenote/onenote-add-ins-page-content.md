---
title: Работа с содержимым страницы в OneNote
description: ''
ms.date: 1/10/2019
ms.openlocfilehash: 617c30f2a9a0c72b1c309ce299f388b5a16b983f
ms.sourcegitcommit: 384e217fd51d73d13ccfa013bfc6e049b66bd98c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/11/2019
ms.locfileid: "27896338"
---
# <a name="work-with-onenote-page-content"></a>Работа с содержимым страницы в OneNote

В API JavaScript для надстроек OneNote содержимое страницы представлено указанной ниже объектной моделью.

  ![Схема объектной модели страницы OneNote](../images/one-note-om-page.png)

- Объект Page содержит коллекцию объектов PageContent.
- Объект PageContent содержит контент типов Outline, Image или Other.
- Объект Outline содержит коллекцию объектов Paragraph.
- Объект Paragraph содержит контент типов RichText, Image, Table или Other.

Чтобы создать пустую страницу OneNote, воспользуйтесь одним из указанных ниже методов.

- [Section.addPage](https://docs.microsoft.com/javascript/api/onenote/onenote.section#addpage-title-)
- [Page.insertPageAsSibling](https://docs.microsoft.com/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

Затем используйте методы в указанных ниже объектах для работы с содержимым страницы, например `Page.addOutline` и `Outline.appendHtml`.

- [Страница](https://docs.microsoft.com/javascript/api/onenote/onenote.page)
- [Структура](https://docs.microsoft.com/javascript/api/onenote/onenote.outline)
- [Абзац](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph)

Для представления содержимого и структуры страницы OneNote используется HTML. Для создания или обновления содержимого страницы поддерживается только подмножество HTML, как описано ниже.

## <a name="supported-html"></a>Поддерживаемые элементы HTML

Для создания и обновления содержимого страницы в API JavaScript для надстроек OneNote используются указанные ниже элементы HTML.

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

Через `Page#load` доступ можно получить только к *содержимому активной страницы*. Чтобы изменить активную страницу, вызовите команду `navigateToPage($page)`.

Метаданные, например "Название", можно запросить для любой страницы.

## <a name="see-also"></a>См. также

- [Обзор API JavaScript для OneNote](onenote-add-ins-programming-overview.md)
- [Справочник по API JavaScript для OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
