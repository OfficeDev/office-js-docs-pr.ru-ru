---
title: Набор обязательных элементов API JavaScript для Word 1,2
description: Сведения о наборе требований WordApi 1,2
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: ee9bf60a3a944a3a01a2ca5aa10d01958e3d3475
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996427"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Новые возможности API JavaScript для Word 1.2

WordApi 1,2 добавлена поддержка встроенных рисунков.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Word 1,2. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых в наборе обязательных элементов API JavaScript для Word 1,2 или более ранней версии, обратитесь к разделам [API Word в наборе требований 1,2](/javascript/api/word?view=word-js-1.2&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет рисунок в содержимое в заданном расположении.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет встроенный рисунок в элемент управления содержимым в указанном расположении.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|Удаляет встроенный рисунок из документа.|
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в указанном расположении.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|Вставляет HTML-код в указанном расположении.|
||[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет встроенный рисунок в указанном расположении.|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|Вставляет текст в заданном расположении.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Возвращает родительский абзац, который содержит встроенный рисунок.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.inlinepicture#select-selectionmode-)|Выбирает встроенный рисунок.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет рисунок в указанном расположении.|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|Возвращает коллекцию объектов встроенных рисунков в диапазоне.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
