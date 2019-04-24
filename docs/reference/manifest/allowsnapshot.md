---
title: Элемент AllowSnapshot в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 02d44167dd1fd46ec6316f3e04393c99f19c9ff0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450676"
---
# <a name="allowsnapshot-element"></a>Элемент AllowSnapshot

Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.

**Тип надстройки:** контентная

## <a name="syntax"></a>Синтаксис

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Примечания

 > [!IMPORTANT]
 > По умолчанию элементу **AllowSnapshot** присвоено значение `true`. Это означает, что пользователи увидят изображение надстройки, если откроют документ в той версии ведущего приложения, которая не поддерживает надстройки Office. Кроме того, если ведущему приложению не удастся подключиться к серверу, на котором размещена надстройка, то отобразится статическое изображение надстройки. Тем не менее, если оставить значение по умолчанию, то возможная конфиденциальная информация в надстройке будет доступна непосредственно из документа, где размещена эта надстройка.

