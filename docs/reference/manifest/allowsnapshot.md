---
title: Элемент AllowSnapshot в файле манифеста
description: Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ea910e1ad747e304dbc6ab4fbdcf44a9610dab19
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294278"
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
 > По умолчанию элементу **AllowSnapshot** присвоено значение `true`. Это делает изображение надстройки видимым для пользователей, открывающих документ в версии приложения Office, не поддерживающей надстройки Office, или предоставляет статическое изображение надстройки, если приложение не может подключиться к серверу, на котором размещается надстройка. Тем не менее, если оставить значение по умолчанию, то возможная конфиденциальная информация в надстройке будет доступна непосредственно из документа, где размещена эта надстройка.
