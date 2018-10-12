---
title: Работа с содержимым страницы в OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 246c864cfb6a63b5f78da8c1189ac5545411168c
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505666"
---
# <a name="work-with-onenote-page-content"></a>Работа с содержимым страницы в OneNote 

В API JavaScript для надстроек OneNote содержимое страницы представлено указанной ниже объектной моделью.

  ![Схема объектной модели страницы OneNote](../images/one-note-om-page.png)

- Объект Page содержит коллекцию объектов PageContent.
- Объект PageContent включает типы содержимого Outline, Image или Other.
- Объект Outline содержит коллекцию объектов Paragraph.
- Объект Paragraph включает типы содержимого  RichText, Image, Table или Other.

Чтобы создать пустую страницу OneNote, воспользуйтесь одним из указанных ниже методов.

- [Section.addPage](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#addpage-title-)
- [Page.insertPageAsSibling](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#insertsectionassibling-location--title-)

Затем используйте методы в указанных ниже объектах для работы с содержимым страницы, например Page.addOutline и Outline.appendHtml. 

- [Страница](https://docs.microsoft.com/javascript/api/onenote/onenote.page?view=office-js)
- [Структура](https://docs.microsoft.com/javascript/api/onenote/onenote.outline?view=office-js)
- [Абзац](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph?view=office-js)

Для представления содержимого и структуры страницы OneNote используется HTML. Для создания или обновления содержимого страницы поддерживается только подмножество HTML, как описано ниже.

## <a name="supported-html"></a>Поддерживаемые элементы HTML

Для создания и обновления содержимого страницы в API JavaScript для надстроек OneNote используются указанные ниже элементы HTML.

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

## <a name="accessing-page-contents"></a>Доступ к содержимому страницы

Через `Page#load` доступ можно получить только к *содержимому активной страницы*. Чтобы изменить активную страницу, вызовите команду `navigateToPage($page)`.

Метаданные, например "Название", можно запросить для любой страницы.

## <a name="see-also"></a>См. также

- [Обзор создания кода с помощью API JavaScript для OneNote](onenote-add-ins-programming-overview.md)
- [Ссылка на API JavaScript для OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
