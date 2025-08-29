# Ten plik przechowuje wszystkie zapytania SQL używane w aplikacji.

sql_queries = {
    #KLUCZOWE WSKAŹNIKI EFEKTYWNOŚCI (KPI) 


    'get_sredni_benefit_na_wniosek': """""",



    #-------------RAPORT OPERACYJNY-----------------

    #1. ZGŁOSZONE
    'get_zgloszone_wnioski': """
        SELECT COUNT(*) 
        FROM dbo.WnioskiIdeaManagement
        WHERE YEAR(DataZgloszenia) = ?
          AND MONTH(DataZgloszenia) BETWEEN 1 AND ?
          AND Dzial = ?
    """,

    #2. ZGŁOSZONE_OTWARTE
    'get_wnioski_otwarte': """
        DECLARE @RokDocelowy INT = ?;
        DECLARE @MiesiacDocelowy INT = ?; 
        DECLARE @Dzial NVARCHAR(100) = ?;

        DECLARE @DataPoczatkowa DATE = DATEFROMPARTS(@RokDocelowy - 1, 1, 1); -- 1 stycznia poprzedniego roku
        DECLARE @DataKoncowa DATE = EOMONTH(DATEFROMPARTS(@RokDocelowy, @MiesiacDocelowy, 1)); -- Ostatni dzień podanego miesiąca

        SELECT
            COUNT(*) AS LiczbaWnioskowOtwartych
        FROM
            dbo.WnioskiIdeaManagement
        WHERE
            Dzial = @Dzial
            AND DataZgloszenia BETWEEN @DataPoczatkowa AND @DataKoncowa
            AND (
                (StatusWniosku = 'Zatwierdzony' AND (ProgressBar <> 1 OR ProgressBar IS NULL))
                OR StatusWniosku IN ('Nowy', 'Do Wyjaśnienia')
            );
    """,

    #3. ZGŁOSZONE_OTWARTE_90DNI
    'get_wnioski_otwarte_90dni': """
          DECLARE @RokDocelowy INT = ?;
          DECLARE @MiesiacDocelowy INT = ?; 
            DECLARE @Dzial NVARCHAR(100) = ?;

            DECLARE @DataPoczatkowa DATE = DATEFROMPARTS(@RokDocelowy - 1, 1, 1); -- 1 stycznia poprzedniego roku
            DECLARE @DataKoncowa DATE = EOMONTH(DATEFROMPARTS(@RokDocelowy, @MiesiacDocelowy, 1)); -- Ostatni dzień podanego miesiąca

            SELECT
                COUNT(*) AS LiczbaWnioskowOtwartych
            FROM
                dbo.WnioskiIdeaManagement
            WHERE
                Dzial = @Dzial
                AND DataZgloszenia BETWEEN @DataPoczatkowa AND @DataKoncowa
                AND (
                    (StatusWniosku = 'Zatwierdzony' AND (ProgressBar <> 1 OR ProgressBar IS NULL))
                    OR StatusWniosku IN ('Nowy', 'Do Wyjaśnienia')
                )
            AND DATEDIFF(DAY, DataZgloszenia, GETDATE()) > 90;
    """,

    #4. ZAAKCEPTOWANE_ZREALIZOWANE
    'get_wnioski_zaakceptowane_zrealizowane': """
        SELECT COUNT(*)
          FROM dbo.WnioskiIdeaManagement
          WHERE Dzial = ?
            AND StatusWniosku = 'Zatwierdzony'
            AND YEAR(DataWdrozenia) = ?
            AND MONTH(DataWdrozenia) <= ?
            AND (DataWdrozenia IS NOT NULL OR DataWdrozenia <> '0001-01-01')
    """,

    #5. ODRZUCONE
    'get_odrzucone': """
       SELECT COUNT(*)
        FROM dbo.WnioskiIdeaManagement
        WHERE StatusWniosku = 'Odrzucony'
          AND Dzial = ?
          AND YEAR(DataZgloszenia) = ?
          AND MONTH(DataZgloszenia) <= ?
    """,

    #6. SREDNI CZAS REALIZACJI
    'get_sredni_czas_realizacji': """
        DECLARE @Dzial NVARCHAR(100) = ?;
        DECLARE @Rok INT = ?;
        DECLARE @Miesiac INT = ?; 

        DECLARE @DataPoczatkowa DATE = DATEFROMPARTS(@Rok - 1 , 1, 1); -- 1 stycznia podanego roku
        DECLARE @DataKoncowa DATE = EOMONTH(DATEFROMPARTS(@Rok, @Miesiac, 1)); -- Ostatni dzień podanego miesiąca

        SELECT
            ISNULL(
                AVG(CAST(DATEDIFF(day, DataZgloszenia, DataWdrozenia) AS INT)),
            0) AS SredniCzasRealizacjiWDniach
        FROM
            dbo.WnioskiIdeaManagement
        WHERE
            StatusWniosku = 'Zatwierdzony'
            AND Dzial = @Dzial
            AND DataWdrozenia BETWEEN @DataPoczatkowa AND @DataKoncowa;
    """, 

    #7. ZYSK WIRTUALNY
    'get_zysk_wirtualny': """
            DECLARE @Rok INT = ?;
            DECLARE @Miesiac INT = ?;
            DECLARE @Dzial NVARCHAR(100) = ?;

            DECLARE @DataPoczatkowa DATE = DATEFROMPARTS(@Rok, 1, 1); -- 1 stycznia podanego roku
            DECLARE @DataKoncowa DATE = EOMONTH(DATEFROMPARTS(@Rok, @Miesiac, 1)); -- Ostatni dzień podanego miesiąca

            SELECT
                ISNULL(
                    SUM(
                        CASE
                            -- Sprawdź warunek dla kolumny 'PoprawaWplynieNa'
                            WHEN PoprawaWplynieNa IN ('BHP&Ergonomia', 'Jakość') THEN
                                -- Jeśli warunek jest spełniony, sprawdź priorytet
                                CASE
                                    WHEN Priorytet = 'Ważny' THEN 500
                                    WHEN Priorytet = 'Wysoki' THEN 1000
                                    ELSE 0 -- Jeśli priorytet jest inny, nie dodawaj nic
                                END
                            ELSE 0 -- Jeśli 'PoprawaWplynieNa' jest inna, nie dodawaj nic
                        END
                    ), 0) AS ObliczonyZyskWirtualny

            FROM
                dbo.WnioskiIdeaManagement
            WHERE
                Dzial = @Dzial
                AND DataZgloszenia BETWEEN @DataPoczatkowa AND @DataKoncowa
                AND StatusWniosku = 'Zatwierdzony'
                AND DataWdrozenia IS NOT NULL;
    """, 

    #8. ZYSK MIERZALNY
    'get_zysk_mierzalny': """
        
    """

    
}