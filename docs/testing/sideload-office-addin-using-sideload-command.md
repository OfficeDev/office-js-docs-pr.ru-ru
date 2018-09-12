---
title: Загрузка неопубликованных надстроек Office с использованием команды sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 1ab0277493f2899adb479c2f24b1635a881af3cc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944043"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Загрузка неопубликованных надстроек Office для тестирования с использованием **команды sideload**
 >[!NOTE]
>Метод "npm run sideload" работает только для надстроек Excel, Word и PowerPoint, которые запускаются в Windows, и только для проектов надстройки, которые были созданы с помощью инструмента [**yo office**](https://github.com/OfficeDev/generator-office)   и у которых есть сценарий `sideload` в разделе `scripts` файла package.json (у проектов, созданных с помощью более ранних версий **yo office** также нет этого сценария). Если ваш проект был создан с помощью Visual Studio или у него нет сценария загрузки неопубликованных приложений, вы можете загрузить его неопубликованным в Windows с помощью метода, описанного в статье [ Загрузка неопубликованной надстройки Office из сетевой папки](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) .
>
> Если вы не тестируете надстройку Word, Excel или PowerPoint в Windows, см. одну из следующих статей.
> 
> - [Загрузка неопубликованных надстроек Office в Office Online для тестирования](sideload-office-add-ins-for-testing.md)
> - [Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [Загрузка неопубликованных надстроек Outlook для тестирования](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. Откройте командную строку от имени администратора.

2. Измените каталоги на корень папки вашего проекта надстройки.

3. Выполните следующую команду, чтобы запустить экземпляр локального сервера в порт 3000 для подачи вашего проекта надстройки: «**npm run start**»

4. Откройте командную строку от имени администратора.

5. Измените каталоги на корневую папку вашего проекта надстройки.

6. Выполните следующую команду для загрузки ведущего приложения (например, Excel, Word) и регистрации надстройки в ведущем приложении: «**npm run sideload**»

## <a name="see-also"></a>См. также

- [Проверка манифеста и устранение связанных с ним неполадок](troubleshoot-manifest.md)
- [Публикация надстройки Office](../publish/publish.md)