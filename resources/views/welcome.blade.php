<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    

    <style>
        body, html {
            margin: 0;
            padding: 0;
            width: 100%;
            height: 100%;
            overflow: hidden;
            font-family: Arial, sans-serif;
            background-color: #f0f0f0;
        }
        spline-viewer {
            width: 100vw;
            height: 100vh;
            display: block;
        }
        .animated-section {
            opacity: 0;
            transform: translateY(50px);
            transition: opacity 0.5s, transform 0.5s;
            padding: 20px;
        }
        .animated-section.in-view {
            opacity: 1;
            transform: translateY(0);
        }
        .animated-text {
            text-align: center;
            margin-top: 20px;
            font-size: 18px;
        }
        .btn-5 {
            position: fixed;
            bottom: 80px;
            left: 50%;
            transform: translateX(-50%);
            width: 130px;
            height: 40px;
            line-height: 42px;
            padding: 0;
            border: none;
            background: rgb(255,27,0);
            background: linear-gradient(0deg, rgba(255,27,0,1) 0%, rgba(251,75,2,1) 100%);
            cursor: pointer;
            color: white;
            text-align: center;
            font-size: 16px;
            position: relative;
            box-sizing: border-box;
            transition: color 0.4s ease, background 0.4s ease;
            border-radius: 20px; /* Köşeleri yumuşatmak için */
        }
        .btn-5:hover {
            color: white;
            background: linear-gradient(0deg, rgba(255,87,34,1) 0%, rgba(251,75,2,1) 100%);
            box-shadow: none;
        }
        .btn-5:before,
        .btn-5:after {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            height: 2px;
            width: 0;
            background: #f0094a;
            box-shadow: -1px -1px 5px 0px #fff, 7px 7px 20px 0px #0003, 4px 4px 5px 0px #0002;
            transition: 400ms ease all;
        }
        .btn-5:after {
            right: inherit;
            top: inherit;
            left: 0;
            bottom: 0;
        }
        .btn-5:hover:before,
        .btn-5:hover:after {
            width: 100%;
            transition: 800ms ease all;
        }
    </style>
</head>
<body>
    <script type="module" src="https://unpkg.com/@splinetool/viewer@1.8.1/build/spline-viewer.js"></script>
    
    <spline-viewer url="https://prod.spline.design/nsINiC0FgNWFrNLf/scene.splinecode"></spline-viewer>

    <script>
        const remRecur = async () => {
            const sleep = async (time) => new Promise((resolve) => setTimeout(() => {
                console.log('I ITERATE')
                resolve()
            }, time));

            let iterCount = 0;
            const MAX_ITER = 50;
            while(iterCount < MAX_ITER) {
                const splineContainer = document.querySelector('spline-viewer');
                if (splineContainer) {
                    const logoEl = splineContainer?.shadowRoot?.querySelector('#logo');
                    if (logoEl) {
                        logoEl.remove();
                        break;
                    }
                    await sleep(10)
                }
                await sleep(10)
            }
        }
        remRecur()
    </script>

    <button class="btn-5" onclick="redirectToLogin()">SZN Portal</button>

    <div class="animated-section">
        <!-- Content of animated section -->
    </div>

    <!-- Other sections with class "animated-section" -->

    <script>
        document.addEventListener('scroll', function() {
            const sections = document.querySelectorAll('.animated-section');
            const triggerBottom = window.innerHeight / 1.3;

            sections.forEach(section => {
                const sectionTop = section.getBoundingClientRect().top;
                if (sectionTop < triggerBottom) {
                    section.classList.add('in-view');
                } else {
                    section.classList.remove('in-view');
                }
            });
        });

        function redirectToLogin() {
            window.location.href = "/index";
        }
    </script>
</body>
</html>
