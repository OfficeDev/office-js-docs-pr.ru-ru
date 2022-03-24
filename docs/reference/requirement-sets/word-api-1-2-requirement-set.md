---
title: Набор API API Word JavaScript 1.2
description: Сведения о наборе требований WordApi 1.2.
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: 1a5af83615786b241c43ecb07ee0d23b3758cfc8
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744225"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Новые возможности API JavaScript для Word 1.2

WordApi 1.2 добавила поддержку для inline pictures.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Word JavaScript, за набором 1.2. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованием API Word JavaScript, установленным 1.2 или ранее, см. в справке Word API в наборе требований [1.2 или ранее](/javascript/api/word?view=word-js-1.2&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertinlinepicturefrombase64-member(1))|Вставляет рисунок в содержимое в заданном расположении.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertinlinepicturefrombase64-member(1))|Вставляет встроенный рисунок в элемент управления содержимым в указанном расположении.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-delete-member(1))|Удаляет встроенный рисунок из документа.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertbreak-member(1))|Вставляет разрыв в указанном расположении в основном документе.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertfilefrombase64-member(1))|Вставляет документ в указанном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserthtml-member(1))|Вставляет HTML-код в указанном расположении.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertinlinepicturefrombase64-member(1))|Вставляет встроенный рисунок в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertooxml-member(1))|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertparagraph-member(1))|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserttext-member(1))|Вставляет текст в заданном расположении.|
||[paragraph](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-paragraph-member)|Возвращает родительский абзац, который содержит встроенный рисунок.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-select-member(1))|Выбирает встроенный рисунок.|
|[Range](/javascript/api/word/word.range)|[inlinePictures](/javascript/api/word/word.range#word-word-range-inlinepictures-member)|Возвращает коллекцию объектов встроенных рисунков в диапазоне.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertinlinepicturefrombase64-member(1))|Вставляет рисунок в указанном расположении.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
