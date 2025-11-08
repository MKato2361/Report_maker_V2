import streamlit as st
import streamlit.components.v1 as components

def inject_pwa_header():
    """PWA用のheadタグを注入"""
    pwa_html = """
    <head>
        <meta name="apple-mobile-web-app-capable" content="yes">
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
        <meta name="apple-mobile-web-app-title" content="故障報告">
        <meta name="mobile-web-app-capable" content="yes">
        <meta name="theme-color" content="#ff4b4b">
        
        <!-- アイコン設定 -->
        <link rel="apple-touch-icon" href="/app/static/icons/icon-192.png">
        <link rel="icon" type="image/png" sizes="192x192" href="/app/static/icons/icon-192.png">
        <link rel="icon" type="image/png" sizes="512x512" href="/app/static/icons/icon-512.png">
        
        <!-- manifest -->
        <link rel="manifest" href="/app/static/manifest.json">
        
        <script>
            if ('serviceWorker' in navigator) {
                window.addEventListener('load', () => {
                    navigator.serviceWorker.register('/app/static/service-worker.js')
                        .then(reg => console.log('Service Worker registered'))
                        .catch(err => console.log('Service Worker registration failed'));
                });
            }
        </script>
    </head>
    """
    components.html(pwa_html, height=0)
