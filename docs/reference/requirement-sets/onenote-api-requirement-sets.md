---
title: Наборы обязательных элементов API JavaScript для OneNote
description: Узнайте больше о наборах обязательных элементов API JavaScript для OneNote.
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: ecdb26edca54758540688ba03b1d9c1eec14e739
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350192"
---
# <a name="onenote-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для OneNote

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

В приведенной ниже таблице перечислены наборы обязательных элементов для OneNote, клиентские приложения Office, которые их поддерживают, а также версии сборок или даты выхода.

|  Набор обязательных элементов  |  Office в Интернете |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1&preserve-view=true)  | Сентябрь 2016 г. |  

## <a name="onenote-javascript-api-11"></a>API JavaScript для OneNote 1.1

API JavaScript для OneNote 1.1 — первая версия этого API. Дополнительные сведения об этом API см. в статье [Обзор создания кода с помощью API JavaScript для OneNote](../../onenote/onenote-add-ins-programming-overview.md).

## <a name="runtime-requirement-support-check"></a>Проверка поддержки обязательных элементов в среде выполнения

В среде выполнения надстройки могут проверять, поддерживает ли конкретное приложение Office набор обязательных элементов API, с помощью следующей проверки.

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Проверка поддержки обязательных элементов в манифесте

Используйте элемент `Requirements` в манифесте надстройки, чтобы указать ключевые наборы обязательных элементов или элементы API, которые должна использовать надстройка. Если платформа или приложение Office не поддерживает наборы обязательных элементов или элементы API, указанные в элементе `Requirements`, надстройка не будет работать в этом приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".

Ниже показана надстройка, которая загружается во всех клиентских приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для OneNote](/javascript/api/onenote)
- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
