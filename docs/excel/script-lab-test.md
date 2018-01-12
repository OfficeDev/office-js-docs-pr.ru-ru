# <a name="testing-script-lab-integration"></a>Проверка интеграции лаборатории скриптов

Это тестовый файл, предназначенный для демонстрации новой функции ScriptLab, которая позволит разработчикам опробовать свои фрагменты в Excel, Word, PowerPoint.  

## <a name="pre-reqs"></a>Требования:
- Вам потребуется URL-адрес представления из фрагмента ScriptLab.
- Примечание. Чтобы опробовать последние фрагменты с помощью ScriptLab, требуется Office 365.  Разработчики могут получить подписку на Office 365 по [специальной программе](https://dev.office.com/devprogram) (исключительно для разработки).  


## <a name="try-it-out-button"></a>Кнопка "Попробовать"
Таким образом мы добавим кнопку "Попробовать", которую рекомендуется связать с фрагментом кода.  Для этого мы используем класс Office UI Fabric, чтобы стилизовать ссылку под кнопку. Для самой ссылки не забудьте задать атрибут *aria label*.

**Демонстрация:**

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Попробовать</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Попробовать</button>


**Код:**
```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## <a name="embed-script-lab-as-an-iframe"></a>Внедрение лаборатории скриптов в виде iframe
В этом режиме мы внедрим фрагмент непосредственно в наши документы в виде iframe. Ширина имеет значение 95 % (от ширины всех остальных фрагментов). Рекомендуем удалить границу рамки iframe.  Высота должна соответствовать высоте фрагмента.

**Демонстрация:**
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

**Код:**
```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## <a name="testing-considerations"></a>Рекомендации по тестированию
Нам нужно проверить мобильные подписки (не на Office 365). Согласно отзывам на сайте office js docs многие разработчики пользовались версией 2013 или более ранней.  

Для пути внедрения нам нужно окончательное утверждение. Кроме того, нам нужно убедиться, что содержимое на странице gist представления соответствует нашим рекомендациям по внедрению специальных возможностей.
