<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Homepage</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/animate.css/4.1.1/animate.min.css"/>
    <style>
        body, html {
            margin: 0;
            padding: 0;
            height: 100%;
            overflow-x: hidden;
        }
        .video-background {
            position: absolute;
            top: 0;
            left: 0;
            height: 100%;
            width: 100%;
            object-fit: cover;
            z-index: -1;
        }
        .content {
            position: relative;
            z-index: 1;
            color: white;
        }
        .section {
            height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            flex-direction: column;
        }
    </style>
</head>
<body>
    <video autoplay muted loop class="video-background">
        <source src="{{URL::asset('assets/videos/background.mp4')}}" type="video/mp4">
    </video>

    <div class="content">
        <div class="section">
            <h1 class="animate__animated animate__fadeInDown">Welcome to Our Website</h1>
            <p class="animate__animated animate__fadeInUp">Scroll down to explore</p>
        </div>
        <div class="section bg-dark">
            <h2 class="animate__animated animate__fadeInLeft">Our Features</h2>
            <p class="animate__animated animate__fadeInRight">Amazing features to blow your mind.</p>
        </div>
        <div class="section bg-secondary">
            <h2 class="animate__animated animate__fadeInLeft">Contact Us</h2>
            <p class="animate__animated animate__fadeInRight">Get in touch with us for more info.</p>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
