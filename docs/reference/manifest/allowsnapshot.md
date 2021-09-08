---
title: Элемент AllowSnapshot в файле манифеста
description: Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ea910e1ad747e304dbc6ab4fbdcf44a9610dab19
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937251"
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
 > По умолчанию элементу **AllowSnapshot** присвоено значение `true`. Это делает изображение надстройки видимым для пользователей, которые открывают документ в версии приложения Office, которое не поддерживает Office надстройки, или обеспечивает статичное изображение надстройки, если приложение не может подключиться к серверу, на который размещена надстройка. Тем не менее, если оставить значение по умолчанию, то возможная конфиденциальная информация в надстройке будет доступна непосредственно из документа, где размещена эта надстройка.
