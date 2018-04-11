---
title: Проверка интеграции Script Lab
description: 'Этот тестовый файл демонстрирует новую функцию ScriptLab, которая позволит разработчикам опробовать свои фрагменты в Excel, Word и PowerPoint.'
ms.date: 03/14/2018
---


# <a name="testing-script-lab-integration"></a>Проверка интеграции Script Lab

Этот тестовый файл демонстрирует новую функцию ScriptLab, которая позволит разработчикам опробовать свои фрагменты в Excel, Word и PowerPoint. 

## <a name="prerequisites"></a>Необходимые компоненты

- Вам потребуется URL-адрес представления из фрагмента ScriptLab.

> [!NOTE] 
> *Следует* отметить, что для изучения последних фрагментов с помощью ScriptLab требуется Office 365. Разработчики могут получить подписку на Office 365 по [специальной программе](https://developer.microsoft.com/en-us/office/dev-program) (исключительно для разработки). Пошаговые инструкции для принятия участия в этой программе, регистрации и настройки подписки см. в [документации по программе для разработчиков приложений для Office 365](https://docs.microsoft.com/ru-ru/office/developer-program/office-365-developer-program). 


## <a name="try-it-out-button"></a>Кнопка "Попробовать"

Так мы добавим кнопку **Попробовать**, которую рекомендуем связать с фрагментом кода. Для этого используем класс Office UI Fabric, чтобы стилизовать ссылку под кнопку. Для самой ссылки не забудьте задать атрибут `aria label`.

### <a name="demo"></a>Демонстрация

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Попробовать</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Попробовать</button>


### <a name="code"></a>Код

```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## <a name="embed-script-lab-as-an-iframe"></a>Внедрение Script Lab в виде iframe

В этом режиме мы внедрим фрагмент непосредственно в наши документы в виде iframe. Ширина имеет значение 95 % (от ширины всех остальных фрагментов). Рекомендуем удалить границу рамки iframe. Высота должна соответствовать фрагменту.

### <a name="demo"></a>Демонстрация

<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

### <a name="code"></a>Код

```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## <a name="testing-considerations"></a>Рекомендации по тестированию

Нам нужно проверить мобильные подписки (не на Office 365). Исходя из отзывов на сайте office-js-docs, многие разработчики пользуются версией 2013 или более ранней.  

Для пути внедрения нам нужно окончательное утверждение. Кроме того, нам нужно убедиться, что содержимое на странице gist представления соответствует нашим рекомендациям по внедрению специальных возможностей.


