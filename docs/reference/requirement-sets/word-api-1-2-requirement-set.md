---
title: Набор API API Word JavaScript 1.2
description: Сведения о наборе требований WordApi 1.2
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: de293cf67bbb452fe3c2b8c5de4896adf5cf7a43
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154823"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Новые возможности API JavaScript для Word 1.2

WordApi 1.2 добавила поддержку для inline pictures.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Word JavaScript, за набором 1.2. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованием API Word JavaScript, установленным 1.2 или ранее, см. в справке Word API в наборе требований [1.2 или более ранних](/javascript/api/word?view=word-js-1.2&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Вставляет рисунок в содержимое в заданном расположении.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Вставляет встроенный рисунок в элемент управления содержимым в указанном расположении.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete__)|Удаляет встроенный рисунок из документа.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertBreak_breakType__insertLocation_)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertFileFromBase64_base64File__insertLocation_)|Вставляет документ в указанном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertHtml_html__insertLocation_)|Вставляет HTML-код в указанном расположении.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Вставляет встроенный рисунок в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertOoxml_ooxml__insertLocation_)|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertParagraph_paragraphText__insertLocation_)|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertText_text__insertLocation_)|Вставляет текст в заданном расположении.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Возвращает родительский абзац, который содержит встроенный рисунок.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#select_selectionMode_)|Выбирает встроенный рисунок.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Вставляет рисунок в указанном расположении.|
||[inlinePictures](/javascript/api/word/word.range#inlinePictures)|Возвращает коллекцию объектов встроенных рисунков в диапазоне.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
