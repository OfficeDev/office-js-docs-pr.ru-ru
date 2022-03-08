> [!NOTE]
> Если вы выполняете надстройки из localhost и видите ошибку "К сожалению, мы не смогли получить доступ *к {your-add-in-name-here}*. Убедитесь, что у вас есть сетевое подключение. Если проблема продолжится, попробуйте еще раз.", возможно, потребуется включить освобождение от циклов.
>
> 1. Закройте Outlook.
> 1. Откройте диспетчер **задач** и убедитесь, что **msoadfsb.exe** процесс не запущен.
> 1. Установите освобождение [от циклов в](/previous-versions/windows/apps/hh780593(v=win.10)?redirectedfrom=MSDN) повышенной подсказке.
>     - Если используется и порт `https://localhost` 3000 (конфигурация по умолчанию), запустите следующую команду.
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_https___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>     - Если используется и порт `http://localhost` 3000, запустите следующую команду.
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>
>      **Примечание**. Если вы не используете порт 3000 по умолчанию, замените его в команде фактическим номером порта.
> 1. Перезапустите Outlook.
