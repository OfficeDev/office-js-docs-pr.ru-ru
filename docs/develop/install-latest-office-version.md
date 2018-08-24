---
title: Установка последней версии Office 2016
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 98dc69a7971a94b96bc3f7304fc7905f31013a87
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925236"
---
# <a name="install-the-latest-version-of-office-2016"></a><span data-ttu-id="f6f86-102">Установка последней версии Office 2016</span><span class="sxs-lookup"><span data-stu-id="f6f86-102">Install the latest version of Office 2016</span></span>

<span data-ttu-id="f6f86-103">Первыми новые функции для разработчиков, в том числе предварительные версии, получают подписчики, которые получают последние сборки Office раньше других.</span><span class="sxs-lookup"><span data-stu-id="f6f86-103">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="f6f86-104">Как получать последние сборки раньше других</span><span class="sxs-lookup"><span data-stu-id="f6f86-104">Opt in to getting the latest builds</span></span>

<span data-ttu-id="f6f86-105">Чтобы получать последние сборки Office 2016 раньше других:</span><span class="sxs-lookup"><span data-stu-id="f6f86-105">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="f6f86-106">Если вы подписаны на Office 365 для дома, Office 365 персональный или Office 365 для студентов, [примите участие в программе предварительной оценки Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="f6f86-106">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="f6f86-107">Если вы пользуетесь Office 365 для бизнеса, прочитайте статью [Установка сборки раннего выпуска для клиентов Office 365 для бизнеса](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="f6f86-107">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="f6f86-108">Если вы используете Office 2016 для Mac:</span><span class="sxs-lookup"><span data-stu-id="f6f86-108">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="f6f86-109">Запустите программу Office 2016 для Mac.</span><span class="sxs-lookup"><span data-stu-id="f6f86-109">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="f6f86-110">Выберите пункт **Проверить наличие обновлений** в меню "Справка".</span><span class="sxs-lookup"><span data-stu-id="f6f86-110">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="f6f86-111">В окне "Автоматическое обновление (Майкрософт)" установите флажок для участия в программе предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="f6f86-111">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="f6f86-112">Как получить последнюю сборку</span><span class="sxs-lookup"><span data-stu-id="f6f86-112">Get the latest build</span></span>

<span data-ttu-id="f6f86-113">Чтобы получить последнюю сборку Office 2016:</span><span class="sxs-lookup"><span data-stu-id="f6f86-113">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="f6f86-114">Скачайте [средство развертывания Office 2016](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="f6f86-114">Download the [Office 2016 Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span> 
2. <span data-ttu-id="f6f86-p101">Запустите это средство. Будут извлечены два файла: Setup.exe и configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="f6f86-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="f6f86-117">Замените файл configuration.xml [файлом конфигурации первого выпуска](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="f6f86-117">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="f6f86-118">Выполните следующую команду от имени администратора: `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="f6f86-118">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="f6f86-119">Команды может выполняться долго, при этом ход ее выполнения нигде не отображается.</span><span class="sxs-lookup"><span data-stu-id="f6f86-119">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="f6f86-p102">По завершении процесса установки у вас будут последние версии приложений Office 2016. Чтобы убедиться, что у вас последняя сборка, в любом приложении Office последовательно выберите **Файл**  >  **Учетная запись**. В разделе "Обновления Office" над номером версии должна быть надпись Office Insiders.</span><span class="sxs-lookup"><span data-stu-id="f6f86-p102">When the installation process finishes, you will have the latest Office 2016 applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Снимок экрана, на котором показаны сведения о продукте с надписью "Участники программы предварительной оценки Office"](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="f6f86-124">Минимальные сборки Office, которые могут использовать наборы обязательных элементов API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="f6f86-124">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="f6f86-125">Сведения о минимальных сборках продуктов для каждой платформы см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="f6f86-125">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="f6f86-126">Наборы обязательных элементов API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="f6f86-126">Word JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="f6f86-127">Наборы обязательных элементов API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="f6f86-127">Excel JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="f6f86-128">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="f6f86-128">OneNote JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="f6f86-129">Наборы обязательных элементов API диалоговых окон</span><span class="sxs-lookup"><span data-stu-id="f6f86-129">Dialog API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="f6f86-130">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="f6f86-130">Office common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
