---
title: Отладка надстроек в Office Online
description: Сведения о том, как тестировать и отлаживать надстройки в Office Online.
ms.date: 03/14/2018
ms.openlocfilehash: fac57e136c07bf33dce62908ea2c12d8be806f7b
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925341"
---
# <a name="debug-add-ins-in-office-online"></a>Отладка надстроек в Office Online


Вы можете создавать надстройки и выполнять их отладку на компьютере, на котором нет Windows или классического клиента Office 2013 или Office 2016 (например, если вы создаете надстройку на компьютере Mac). В этой статье рассказывается, как использовать Office Online для тестирования и отладки надстроек. 

## <a name="prerequisites"></a>Необходимые компоненты

Чтобы приступить к работе, выполните указанные ниже действия.

- Получите учетную запись разработчика приложений для Office 365 (если у вас еще нет ее) или доступ к сайту SharePoint.
    
  > [!NOTE]
  > Чтобы бесплатно получить подписку разработчика приложений для Office 365, примите участие в нашей [программе для разработчиков приложений Office 365](https://developer.microsoft.com/office/dev-program). Пошаговые инструкции для принятия участия в этой программе, регистрации и настройки подписки см. в [документации по программе для разработчиков приложений для Office 365](https://docs.microsoft.com/office/developer-program/office-365-developer-program).
     
- Настройте каталог надстроек в Office 365 (SharePoint Online). Каталог надстроек — это специальное семейство веб-сайтов в SharePoint Online, в котором размещены библиотеки документов для надстроек Office. Если у вас есть сайт SharePoint, вы можете настроить библиотеку документов каталога надстроек. Дополнительные сведения см. в статье [Публикация надстроек области задач и контентных надстроек в каталоге надстроек в SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>Отладка надстройки в Excel Online и Word Online

Для отладки надстройки с помощью Office Online выполните указанные ниже действия.

1. Разверните надстройку на сервере, поддерживающем SSL.
    
    > [!NOTE]
    > Рекомендуем использовать [генератор Yeoman](https://github.com/OfficeDev/generator-office) для создания и размещения надстройки.
     
2. В [файле манифеста надстройки](../develop/add-in-manifests.md) измените значение элемента **SourceLocation** так, чтобы оно включало абсолютный URL-адрес, а не относительный. Пример:
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. Выложите манифест в библиотеку надстроек Office в каталоге надстроек в SharePoint.
    
4. В Office 365 в средстве запуска приложений запустите Excel Online или Word Online и откройте новый документ.
    
5. Чтобы добавить вашу надстройку и протестировать ее в приложении, на вкладке "Вставка" выберите **Мои надстройки** или **Надстройки Office**.
    
6. Выполните отладку надстройки в удобном для вас браузерном отладчике.

## <a name="potential-issues"></a>Возможные проблемы    

Ниже указаны некоторые проблемы, которые могут возникнуть при отладке.
    
- Причиной некоторых отображаемых ошибок JavaScript может быть Office Online.
      
- Браузер может отобразить сообщение об ошибке, связанной с недопустимым сертификатом, которое необходимо обойти.
      
- Если вы задаете точки останова в коде, Office Online может отобразить сообщение об ошибке, свидетельствующее о том, что не удается выполнить сохранение.

## <a name="see-also"></a>См. также

- [Рекомендации по разработке надстроек Office](../concepts/add-in-development-best-practices.md)
- 
  [Политики проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies)  
- 
  [Создание эффективных приложений и надстроек AppSource](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)  
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](testing-and-troubleshooting.md)
    
