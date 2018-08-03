---
title: Загрузка неопубликованных надстроек Office с использованием команды sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: c3b53a70b5696e422653350de18d99be16d1d597
ms.sourcegitcommit: 0d4d78e275249f0d4b6a6cf807b42b79890c3023
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2018
ms.locfileid: "21773596"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="baead-102">Загрузка неопубликованных надстроек Office для тестирования с использованием **команды sideload**</span><span class="sxs-lookup"><span data-stu-id="baead-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="baead-103">Метод "npm run sideload" работает только для надстроек Excel, Word и PowerPoint, которые запускаются в Windows, и только для проектов надстройки, которые были созданы с помощью [**инструмента**yo office](https://github.com/OfficeDev/generator-office) и которые имеют `sideload` сценарий в `scripts` разделе файла package.json.</span><span class="sxs-lookup"><span data-stu-id="baead-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="baead-104">(Проекты, созданные со старыми версиями **yo office**, также не имеют этого сценария.) Если ваш проект был создан с помощью Visual Studio или не имеет сценария загрузки неопубликованных приложений, вы можете загрузить его неопубликованным в Windows с помощью метода, описанного в статье [Загрузка неопубликованной надстройки Office из сетевой папки](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="baead-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="baead-105">Если вы не тестируете надстройку Word, Excel или PowerPoint в Windows, см. одну из следующих статей, чтобы загрузить ваше неопубликованное приложение:</span><span class="sxs-lookup"><span data-stu-id="baead-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="baead-106">Загрузка неопубликованных надстроек Office в Office Online для тестирования</span><span class="sxs-lookup"><span data-stu-id="baead-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="baead-107">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="baead-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="baead-108">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="baead-108">Sideload Outlook add-ins for testing</span></span>](../../../../outlook/add-insSideload Outlook Add-ins for testing)

1. <span data-ttu-id="baead-109">Откройте командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="baead-109">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="baead-110">Измените каталоги на корень папки вашего проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="baead-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="baead-111">Выполните следующую команду, чтобы запустить экземпляр локального сервера в порт 3000 для подачи вашего проекта надстройки: «**npm run start**»</span><span class="sxs-lookup"><span data-stu-id="baead-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="baead-112">Откройте командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="baead-112">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="baead-113">Измените каталоги на корневую папку вашего проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="baead-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="baead-114">Выполните следующую команду для загрузки ведущего приложения (например, Excel, Word) и регистрации надстройки в ведущем приложении: «**npm run sideload**»</span><span class="sxs-lookup"><span data-stu-id="baead-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="baead-115">См. также</span><span class="sxs-lookup"><span data-stu-id="baead-115">See also</span></span>

- [<span data-ttu-id="baead-116">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="baead-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="baead-117">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="baead-117">Publish your Office Add-in</span></span>](../publish/publish.md)