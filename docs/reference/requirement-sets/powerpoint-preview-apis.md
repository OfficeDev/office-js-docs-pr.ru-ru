---
title: API предварительного просмотра JavaScript для PowerPoint
description: Сведения о предстоящих API JavaScript для PowerPoint.
ms.date: 11/09/2020
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: b53b6638b16b2028342003b9a77aa59e7406d5f3
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996524"
---
# <a name="powerpoint-javascript-preview-apis"></a>API предварительного просмотра JavaScript для PowerPoint

Новые API JavaScript для PowerPoint впервые представлены в "предварительной версии", а потом — в определенном наборе обязательных элементов после тестирования и получении отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Вставка и удаление слайдов | Позволяет вставлять существующие слайды в текущую презентацию из другой презентации, а также возможность удалять силдес. | [Слайд. Delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript для PowerPoint, находящиеся в предварительной версии. Полный список всех API JavaScript для PowerPoint (в том числе API предварительного просмотра и ранее выпущенных API) представлен в статье [все API JavaScript для PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[Formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Задает формат, используемый при вставке слайдов.|
||[саурцеслидеидс](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Указывает слайды из исходной презентации, которые будут вставлены в текущую презентацию.|
||[таржетслидеид](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Указывает, где будут вставляться новые слайды в презентации.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64File: строка, параметры?: PowerPoint. Инсертслидеоптионс)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Вставляет указанные слайды из презентации в текущую презентацию.|
||[Титуль](/javascript/api/powerpoint/powerpoint.presentation#slides)|Возвращает упорядоченную коллекцию слайдов в презентации.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Удаляет слайд из презентации.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Получает уникальный идентификатор слайда.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Получает количество слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Получает слайд с помощью уникального идентификатора.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Получает слайд с использованием индекса, основанного на нуле, в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Получает слайд с помощью уникального идентификатора.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)
