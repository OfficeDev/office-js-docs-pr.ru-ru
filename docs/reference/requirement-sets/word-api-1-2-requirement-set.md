---
title: Набор обязательных элементов API JavaScript для Word 1,2
description: Сведения о наборе требований WordApi 1,2
ms.date: 07/17/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: c6244b7ce9ff7b5cbde68baad26e60a6326199d8
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804708"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Новые возможности API JavaScript для Word 1.2

WordApi 1,2 добавлена поддержка встроенных рисунков.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API, добавленные в состав набора требований WordApi 1,2.

| Класс | Поля | Описание |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет рисунок в содержимое в заданном расположении. Возможные значения insertLocation: Start и End.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет встроенный рисунок в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|Удаляет встроенный рисунок из документа.|
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе. Возможные значения InsertLocation: Before и After.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в указанном расположении. Возможные значения InsertLocation: Before и After.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|Вставляет HTML-код в указанном расположении. Возможные значения InsertLocation: Before и After.|
||[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет встроенный рисунок в указанном расположении. Значение insertLocation может быть "Replace", "Before" или "After".|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|Вставляет OOXML-код в указанном расположении.  Возможные значения InsertLocation: Before и After.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении. Возможные значения InsertLocation: Before и After.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|Вставляет текст в заданном расположении. Возможные значения insertLocation: Before и After.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Возвращает родительский абзац, который содержит встроенный рисунок. Только для чтения.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.inlinepicture#select-selectionmode-)|Выбирает встроенный рисунок. При этом Word переходит к выделенному объекту.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет рисунок в указанном расположении. Значение insertLocation может быть "Replace", "Start", "End", "Before" или "After".|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|Возвращает коллекцию объектов встроенных рисунков в диапазоне. Только для чтения.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
