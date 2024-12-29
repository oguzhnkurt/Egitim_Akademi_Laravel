<!DOCTYPE html>
<html lang="tr">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sınav Sonucunuz</title>
    <style>
        body {
            background: #111115;
            display: flex;
            justify-content: center;
            align-items: center;
            flex-direction: column;
            height: 100vh;
            color: white;
            font-family: Arial, sans-serif;
        }

        .heart {
            font-size: 6em;
            position: relative;
            margin-bottom: 20px;
        }

        .heartbeat {
            position: relative;
            z-index: 1;
            animation: beat 2s linear infinite;
        }

        .heartecho {
            position: absolute;
            top: 0;
            right: 0;
            bottom: 0;
            left: 0;
            animation: echo 2s linear infinite;
        }

        @keyframes beat {
            0% {
                transform: scale(1);
            }

            14% {
                transform: scale(0.9);
            }

            21% {
                transform: scale(1.1) skew(0.004turn);
            }

            28% {
                transform: scale(1) skew(0.008turn);
            }

            35% {
                transform: scale(1) skew(0);
            }
        }

        @keyframes echo {
            0% {
                opacity: 0.5;
                transform: scale(1);
            }

            14% {
                opacity: 0.4;
                transform: scale(0.8);
            }

            21% {
                opacity: 0.4;
                transform: scale(1.1);
            }

            100% {
                opacity: 0;
                transform: scale(3);
            }
        }

        .result {
            font-size: 2em;
            text-align: center;
            margin-bottom: 40px;
        }

        /* Parlayan Kırmızı Buton */
        .button-85 {
            padding: 0.6em 2em;
            border: none;
            outline: none;
            color: rgb(255, 255, 255);
            background: #111;
            cursor: pointer;
            position: relative;
            z-index: 0;
            border-radius: 10px;
            user-select: none;
            -webkit-user-select: none;
            touch-action: manipulation;
            font-size: 1.2em;
            transition: background-color 0.3s ease;
        }

        .button-85:hover {
            background-color: #ff0000;
        }

        .button-85:before {
            content: "";
            background: linear-gradient(45deg,
                    #ff0000,
                    #ff0000);
            position: absolute;
            top: -2px;
            left: -2px;
            background-size: 400%;
            z-index: -1;
            filter: blur(8px);
            -webkit-filter: blur(8px);
            width: calc(100% + 4px);
            height: calc(100% + 4px);
            animation: glowing-button-85 5s linear infinite;
            transition: opacity 0.3s ease-in-out;
            border-radius: 10px;
        }

        @keyframes glowing-button-85 {
            0% {
                background-position: 0 0;
                opacity: 0.6;
            }

            50% {
                background-position: 200% 0;
                opacity: 1;
            }

            100% {
                background-position: 0 0;
                opacity: 0.6;
            }
        }

        .button-85:after {
            z-index: -1;
            content: "";
            position: absolute;
            width: 100%;
            height: 100%;
            background: #222;
            left: 0;
            top: 0;
            border-radius: 10px;
        }
    </style>
</head>

<body>

    <div class="heart">
        <div class="heartbeat">✏️📑</div>
        <div class="heartecho">✏️📑</div>
    </div>

    <div class="result">
        <p>Sınav Sonucunuz</p>
        <p>Puanınız: {{ $examResult->score }}</p> <!-- Veritabanından gelen sınav sonucu burada -->
    </div>

    <!-- Ana Sayfaya Dön Butonu -->
    <a href="{{ url('/index') }}" class="button-85" role="button">Ana Sayfaya Dön</a>
</body>

</html>