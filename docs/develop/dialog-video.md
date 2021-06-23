---
title: Используйте диалоговое окно "Office" для воспроизведения видео
description: Узнайте, как открыть и сыграть видео в диалоговом Office диалоговом окне
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: bc69827164f2e5a2fed03239566ff814db0397b9
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076071"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="cfa1e-103">Чтобы показать видео, Office диалоговое окно</span><span class="sxs-lookup"><span data-stu-id="cfa1e-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="cfa1e-104">В этой статье рассказывается, как играть видео в диалоговом окне Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="cfa1e-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="cfa1e-105">В этой статье предполагается, что вы знакомы с основами использования диалогового окна Office, как описано в статье [Использование API](dialog-api-in-office-add-ins.md)диалоговых Office в Office надстройки .</span><span class="sxs-lookup"><span data-stu-id="cfa1e-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="cfa1e-106">Чтобы играть видео в диалоговом окне с API Office диалоговом окне, выполните следующие действия:</span><span class="sxs-lookup"><span data-stu-id="cfa1e-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="cfa1e-107">Создайте страницу, содержащую iframe и отсутствие другого контента.</span><span class="sxs-lookup"><span data-stu-id="cfa1e-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="cfa1e-108">Страница должна быть в том же домене, что и хост-страница.</span><span class="sxs-lookup"><span data-stu-id="cfa1e-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="cfa1e-109">Напоминая о том, что такое хост-страница, см. в странице Откройте диалоговое [окно с хост-страницы.](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)</span><span class="sxs-lookup"><span data-stu-id="cfa1e-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="cfa1e-110">В `src` атрибуте iframe указать URL-адрес онлайн-видео.</span><span class="sxs-lookup"><span data-stu-id="cfa1e-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="cfa1e-111">URL-адрес видео должен быть защищен с помощью протокола HTTPS.</span><span class="sxs-lookup"><span data-stu-id="cfa1e-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="cfa1e-112">В этой статье мы назовем эту страницу "video.dialogbox.html".</span><span class="sxs-lookup"><span data-stu-id="cfa1e-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="cfa1e-113">Ниже приведен пример разметки:</span><span class="sxs-lookup"><span data-stu-id="cfa1e-113">The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="cfa1e-114">Используйте вызов `displayDialogAsync` на главной странице, чтобы открыть страницу video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="cfa1e-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="cfa1e-115">Если надстройке необходимо знать, когда пользователь закрывает диалоговое окно, зарегистрируйте обработок события и обработите событие `DialogEventReceived` 12006.</span><span class="sxs-lookup"><span data-stu-id="cfa1e-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="cfa1e-116">Подробные сведения см. в материале [Errors and events in the Office диалоговом окне](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="cfa1e-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="cfa1e-117">Пример видео, играемого в диалоговом окне, см. в примере шаблона дизайна [видео-placemat.](../design/first-run-experience-patterns.md#video-placemat)</span><span class="sxs-lookup"><span data-stu-id="cfa1e-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![Снимок экрана, показывающий воспроизведение видео в диалоговом окне надстройки перед Excel.](../images/video-placemats-dialog-open.png)
