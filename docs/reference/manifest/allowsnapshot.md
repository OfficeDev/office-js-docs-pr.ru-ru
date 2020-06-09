---
title: Элемент AllowSnapshot в файле манифеста
description: Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c46dcd882592c0b015dae4b9774533b96fe75cfe
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608791"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="f7219-103">Элемент AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="f7219-103">AllowSnapshot element</span></span>

<span data-ttu-id="f7219-104">Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.</span><span class="sxs-lookup"><span data-stu-id="f7219-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="f7219-105">**Тип надстройки:** контентная</span><span class="sxs-lookup"><span data-stu-id="f7219-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="f7219-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="f7219-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="f7219-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="f7219-107">Contained in</span></span>

[<span data-ttu-id="f7219-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f7219-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="f7219-109">Примечания</span><span class="sxs-lookup"><span data-stu-id="f7219-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="f7219-110">По умолчанию элементу **AllowSnapshot** присвоено значение `true`.</span><span class="sxs-lookup"><span data-stu-id="f7219-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="f7219-111">Это означает, что пользователи увидят изображение надстройки, если откроют документ в той версии ведущего приложения, которая не поддерживает надстройки Office. Кроме того, если ведущему приложению не удастся подключиться к серверу, на котором размещена надстройка, то отобразится статическое изображение надстройки.</span><span class="sxs-lookup"><span data-stu-id="f7219-111">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="f7219-112">Тем не менее, если оставить значение по умолчанию, то возможная конфиденциальная информация в надстройке будет доступна непосредственно из документа, где размещена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="f7219-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

