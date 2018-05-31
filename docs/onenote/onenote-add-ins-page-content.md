---
title: Работа с содержимым страницы в OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d05f251a798a7670983187bfa4c80140b30f6147
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438860"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="2c01b-102">Работа с содержимым страницы в OneNote</span><span class="sxs-lookup"><span data-stu-id="2c01b-102">Work with OneNote page content</span></span> 

<span data-ttu-id="2c01b-103">В API JavaScript для надстроек OneNote содержимое страницы представлено указанной ниже объектной моделью.</span><span class="sxs-lookup"><span data-stu-id="2c01b-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Схема объектной модели страницы OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="2c01b-105">Объект Page содержит коллекцию объектов PageContent.</span><span class="sxs-lookup"><span data-stu-id="2c01b-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="2c01b-106">Объект PageContent содержит контент типов Outline, Image или Other.</span><span class="sxs-lookup"><span data-stu-id="2c01b-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="2c01b-107">Объект Outline содержит коллекцию объектов Paragraph.</span><span class="sxs-lookup"><span data-stu-id="2c01b-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="2c01b-108">Объект Paragraph содержит контент типов RichText, Image, Table или Other.</span><span class="sxs-lookup"><span data-stu-id="2c01b-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="2c01b-109">Чтобы создать пустую страницу OneNote, воспользуйтесь одним из указанных ниже методов.</span><span class="sxs-lookup"><span data-stu-id="2c01b-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="2c01b-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="2c01b-110">Section.addPage</span></span>](https://dev.office.com/reference/add-ins/onenote/section#addpagetitle-string)
- [<span data-ttu-id="2c01b-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="2c01b-111">Page.insertPageAsSibling</span></span>](https://dev.office.com/reference/add-ins/onenote/page#insertpageassiblinglocation-string-title-string)

<span data-ttu-id="2c01b-112">Затем используйте методы в указанных ниже объектах для работы с содержимым страницы, например Page.addOutline и Outline.appendHtml.</span><span class="sxs-lookup"><span data-stu-id="2c01b-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="2c01b-113">Страница</span><span class="sxs-lookup"><span data-stu-id="2c01b-113">Page</span></span>](https://dev.office.com/reference/add-ins/onenote/page)
- [<span data-ttu-id="2c01b-114">Структура</span><span class="sxs-lookup"><span data-stu-id="2c01b-114">Outline</span></span>](https://dev.office.com/reference/add-ins/onenote/outline)
- [<span data-ttu-id="2c01b-115">Абзац</span><span class="sxs-lookup"><span data-stu-id="2c01b-115">Paragraph</span></span>](https://dev.office.com/reference/add-ins/onenote/paragraph)

<span data-ttu-id="2c01b-p101">Для представления содержимого и структуры страницы OneNote используется HTML. Для создания или обновления содержимого страницы поддерживается только подмножество HTML, как описано ниже.</span><span class="sxs-lookup"><span data-stu-id="2c01b-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="2c01b-118">Поддерживаемые элементы HTML</span><span class="sxs-lookup"><span data-stu-id="2c01b-118">Supported HTML</span></span>

<span data-ttu-id="2c01b-119">Для создания и обновления содержимого страницы в API JavaScript для надстроек OneNote используются указанные ниже элементы HTML.</span><span class="sxs-lookup"><span data-stu-id="2c01b-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="2c01b-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="2c01b-120"></span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="2c01b-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="2c01b-121"></span></span> 
- <span data-ttu-id="2c01b-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="2c01b-122"></span></span>
- <span data-ttu-id="2c01b-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="2c01b-123"></span></span>
- <span data-ttu-id="2c01b-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="2c01b-124"></span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="2c01b-125">Доступ к содержимому страницы</span><span class="sxs-lookup"><span data-stu-id="2c01b-125">Accessing page contents</span></span>

<span data-ttu-id="2c01b-p102">Через `Page#load` доступ можно получить только к *содержимому активной страницы*. Чтобы изменить активную страницу, вызовите команду `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="2c01b-p102">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="2c01b-128">Метаданные, например "Название", можно запросить для любой страницы.</span><span class="sxs-lookup"><span data-stu-id="2c01b-128">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="2c01b-129">См. также</span><span class="sxs-lookup"><span data-stu-id="2c01b-129">See also</span></span>

- [<span data-ttu-id="2c01b-130">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="2c01b-130">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="2c01b-131">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="2c01b-131">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="2c01b-132">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="2c01b-132">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="2c01b-133">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="2c01b-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
