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
# <a name="allowsnapshot-element"></a><span data-ttu-id="94ce4-103">Элемент AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="94ce4-103">AllowSnapshot element</span></span>

<span data-ttu-id="94ce4-104">Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.</span><span class="sxs-lookup"><span data-stu-id="94ce4-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="94ce4-105">**Тип надстройки:** контентная</span><span class="sxs-lookup"><span data-stu-id="94ce4-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="94ce4-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="94ce4-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="94ce4-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="94ce4-107">Contained in</span></span>

[<span data-ttu-id="94ce4-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="94ce4-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="94ce4-109">Примечания</span><span class="sxs-lookup"><span data-stu-id="94ce4-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="94ce4-110">По умолчанию элементу **AllowSnapshot** присвоено значение `true`.</span><span class="sxs-lookup"><span data-stu-id="94ce4-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="94ce4-111">Это делает изображение надстройки видимым для пользователей, открывающих документ в версии приложения Office, не поддерживающей надстройки Office, или предоставляет статическое изображение надстройки, если приложение не может подключиться к серверу, на котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="94ce4-111">This makes an image of the add-in visible for users that open the document in a version of the Office application that doesn't support Office Add-ins, or provides a static image of the add-in if the application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="94ce4-112">Тем не менее, если оставить значение по умолчанию, то возможная конфиденциальная информация в надстройке будет доступна непосредственно из документа, где размещена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="94ce4-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>
