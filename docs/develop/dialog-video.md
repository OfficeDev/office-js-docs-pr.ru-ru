---
title: Используйте диалоговое окно "Office" для воспроизведения видео
description: Сведения о том, как открыть и прослушать видео в диалоговом окне Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 9c65dfb9c0cf1adbc827be25b655e380dc39e2d2
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596531"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="67a38-103">Отображение видео с помощью диалогового окна Office</span><span class="sxs-lookup"><span data-stu-id="67a38-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="67a38-104">В этой статье объясняется, как воссоздать видео в диалоговом окне надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="67a38-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="67a38-105">В этой статье предполагается, что вы знакомы с основами использования диалогового окна Office, как описано в статье [Использование API диалоговых окон Office в](dialog-api-in-office-add-ins.md)надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="67a38-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="67a38-106">Для проигрывания видео в диалоговом окне с помощью API диалогового окна Office выполните следующие действия:</span><span class="sxs-lookup"><span data-stu-id="67a38-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="67a38-107">Создание страницы, содержащей IFRAME, без другого контента.</span><span class="sxs-lookup"><span data-stu-id="67a38-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="67a38-108">Страница должна находиться в том же домене, что и Главная страница.</span><span class="sxs-lookup"><span data-stu-id="67a38-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="67a38-109">Напоминание о странице ведущего приложения можно узнать в разделе [Открытие диалогового окна на странице узла](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span><span class="sxs-lookup"><span data-stu-id="67a38-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="67a38-110">В `src` атрибуте IFRAME укажите URL-адрес видео в Интернете.</span><span class="sxs-lookup"><span data-stu-id="67a38-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="67a38-111">URL-адрес видео должен быть защищен с помощью протокола HTTPS.</span><span class="sxs-lookup"><span data-stu-id="67a38-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="67a38-112">В этой статье мы вызываем эту страницу "Video. DialogBox. HTML".</span><span class="sxs-lookup"><span data-stu-id="67a38-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="67a38-113">Ниже приведен пример разметки:</span><span class="sxs-lookup"><span data-stu-id="67a38-113">The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="67a38-114">Используйте вызов `displayDialogAsync` на главной странице, чтобы открыть страницу video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="67a38-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="67a38-115">Если надстройка должна знать, когда пользователь закрывает диалоговое окно, зарегистрируйте обработчик для `DialogEventReceived` события и обработайте событие 12006.</span><span class="sxs-lookup"><span data-stu-id="67a38-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="67a38-116">Дополнительные сведения: ["ошибки и события" в диалоговом окне Office](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="67a38-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="67a38-117">Пример видеоконференций, воспроизводимого в диалоговом окне, приведен в статье [Образец оформления видео представление](../design/first-run-experience-patterns.md#video-placemat).</span><span class="sxs-lookup"><span data-stu-id="67a38-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![Снимок экрана: диалоговое окно воспроизведения видео в диалоговом окне надстройки](../images/video-placemats-dialog-open.png)
