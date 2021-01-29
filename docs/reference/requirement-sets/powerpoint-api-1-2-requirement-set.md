---
title: Набор требований API JavaScript для PowerPoint 1.2
description: Сведения о наборе требований PowerPointApi 1.2.
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 0aa82b8edc6aab65ebcce7c6bfcb50471c9e38e9
ms.sourcegitcommit: 3123b9819c5225ee45a5312f64be79e46cbd0e3c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/29/2021
ms.locfileid: "50043878"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>Новые возможности API JavaScript для PowerPoint 1.2

В PowerPointApi 1.2 добавлена поддержка вставки слайдов из другой презентации в текущую презентацию и удаления слайдов.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Вставка и удаление слайдов](../../powerpoint/insert-slides-into-presentation.md) | Позволяет вставлять существующие слайды в текущую презентацию из другой презентации, а также удалять слайды. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены наборы требований API JavaScript для PowerPoint 1.2. Полный список всех API JavaScript для PowerPoint (включая API предварительной версии и ранее выпущенные API) см. во всех API [JavaScript для PowerPoint.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Указывает форматирование, которое будет применяться во время вставки слайда.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Указывает слайды из презентации источника, которые будут вставлены в текущую презентацию.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Указывает место вставки новых слайдов в презентацию.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Вставляет указанные слайды из презентации в текущую презентацию.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Возвращает упорядоченную коллекцию слайдов в презентации.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Удаляет слайд из презентации.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Получает уникальный ИД слайда.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Получает количество слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Получает слайд с использованием уникального ИД.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Получает слайд с помощью индекса на основе нуля в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Получает слайд с использованием уникального ИД.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)
