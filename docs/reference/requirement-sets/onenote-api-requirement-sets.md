---
title: Наборы обязательных элементов API JavaScript для OneNote
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 287e405955477a98854b1df4a81fe90ec16e5bbc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450102"
---
# <a name="onenote-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для OneNote

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

В приведенной ниже таблице перечислены наборы обязательных элементов для OneNote, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.

|  Набор обязательных элементов  |  Office Online | 
|:-----|:-----|
| OneNoteApi 1.1  | Сентябрь 2016 г. |  

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="onenote-javascript-api-11"></a>API JavaScript для OneNote 1.1 

API JavaScript для OneNote 1.1 — первая версия этого API. Дополнительные сведения об этом API см. в статье [Обзор создания кода с помощью API JavaScript для OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).

## <a name="runtime-requirement-support-check"></a>Проверка поддержки требований в среде выполнения

Во время выполнения кода надстройки могут проверять, поддерживает ли ведущее приложение набор обязательных элементов API, выполняя следующую проверку: 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a>Проверка поддержки обязательных элементов в манифесте

Используйте элемент Requirements в манифесте надстройки, чтобы указать ключевые наборы требований или элементы API, которые должна использовать надстройка. Если платформа или ведущее приложение Office не поддерживает наборы требований или элементы API, указанные в элементе Requirements, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".

Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
