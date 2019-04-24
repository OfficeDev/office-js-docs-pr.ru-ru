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
# <a name="allowsnapshot-element"></a><span data-ttu-id="2b105-102">Элемент AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="2b105-102">AllowSnapshot element</span></span>

<span data-ttu-id="2b105-103">Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.</span><span class="sxs-lookup"><span data-stu-id="2b105-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="2b105-104">**Тип надстройки:** контентная</span><span class="sxs-lookup"><span data-stu-id="2b105-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="2b105-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2b105-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="2b105-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2b105-106">Contained in</span></span>

[<span data-ttu-id="2b105-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="2b105-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="2b105-108">Примечания</span><span class="sxs-lookup"><span data-stu-id="2b105-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="2b105-109">По умолчанию элементу **AllowSnapshot** присвоено значение `true`.</span><span class="sxs-lookup"><span data-stu-id="2b105-109">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="2b105-110">Это означает, что пользователи увидят изображение надстройки, если откроют документ в той версии ведущего приложения, которая не поддерживает надстройки Office. Кроме того, если ведущему приложению не удастся подключиться к серверу, на котором размещена надстройка, то отобразится статическое изображение надстройки.</span><span class="sxs-lookup"><span data-stu-id="2b105-110">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="2b105-111">Тем не менее, если оставить значение по умолчанию, то возможная конфиденциальная информация в надстройке будет доступна непосредственно из документа, где размещена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="2b105-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

