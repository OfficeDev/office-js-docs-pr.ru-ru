---
title: Загрузка неопубликованных надстроек Office с использованием команды sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 3aacfdb09f362ea10ba0e2393caca335fe4c04c6
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925103"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Загрузка неопубликованных надстроек Office для тестирования с использованием **команды sideload**
 >[!NOTE]
>Метод "npm run sideload" работает только для надстроек Excel, Word и PowerPoint, которые запускаются в Windows, и только для проектов надстройки, которые были созданы с помощью [**инструмента**yo office](https://github.com/OfficeDev/generator-office) и которые имеют `sideload` сценарий в `scripts` разделе файла package.json. (Проекты, созданные со старыми версиями **yo office**, также не имеют этого сценария.) Если ваш проект был создан с помощью Visual Studio или не имеет сценария загрузки неопубликованных приложений, вы можете загрузить его неопубликованным в Windows с помощью метода, описанного в статье [Загрузка неопубликованной надстройки Office из сетевой папки](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
>
> Если вы не тестируете надстройку Word, Excel или PowerPoint в Windows, см. одну из следующих статей, чтобы загрузить ваше неопубликованное приложение:
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