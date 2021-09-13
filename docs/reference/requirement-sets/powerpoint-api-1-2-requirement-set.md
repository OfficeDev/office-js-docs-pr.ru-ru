---
title: PowerPoint Требования к API JavaScript 1.2
description: Сведения о наборе требований PowerPointApi 1.2.
ms.date: 01/27/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: b62bed8d28eb2bacff0450e749da8cf69c868e38
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154314"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>Новые возможности в PowerPoint API JavaScript 1.2

PowerPointApi 1.2 добавила поддержку для вставки слайдов из другой презентации в текущую презентацию и удаления слайдов.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Вставка и удаление слайдов](../../powerpoint/insert-slides-into-presentation.md) | Позволяет вставлять существующие слайды в текущую презентацию из другой презентации, а также возможность удаления слайдов. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены PowerPoint API JavaScript, установленный 1.2. Полный список всех API PowerPoint JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. PowerPoint [API JavaScript.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[форматирование](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Указывает форматирование, которое необходимо использовать во время вставки слайда.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceSlideIds)|Указывает слайды из исходных презентаций, которые будут вставлены в текущую презентацию.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetSlideId)|Указывает, где в презентации будут вставлены новые слайды.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)|Вставляет указанные слайды из презентации в текущую презентацию.|
||[слайды](/javascript/api/powerpoint/powerpoint.presentation#slides)|Возвращает упорядоченную коллекцию слайдов в презентации.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete__)|Удаляет слайд из презентации.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Получает уникальный ID слайда.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getCount__)|Получает количество слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItem_key_)|Получает слайд с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)|Получает слайд с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemOrNullObject_id_)|Получает слайд с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>Дополнительные материалы

- [PowerPoint Справочная документация по API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)
