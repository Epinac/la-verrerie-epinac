/* Domaine de la Verrerie — Site JS
   Vanilla, durable. Une seule source de vérité pour le comportement du site. */
(function () {
    'use strict';

    /* ---- 1. Topbar scroll state ---- */
    function setupScrollState() {
        var threshold = 60, scrolled = false;
        window.addEventListener('scroll', function () {
            var s = window.scrollY > threshold;
            if (s !== scrolled) {
                document.body.classList.toggle('is-scrolled', s);
                scrolled = s;
            }
        }, { passive: true });
    }

    /* ---- 2. Mobile burger toggle ---- */
    function setupBurger() {
        var burger = document.querySelector('.topbar-burger');
        if (!burger) return;
        burger.addEventListener('click', function (e) {
            e.preventDefault();
            var open = document.body.classList.toggle('is-nav-open');
            burger.setAttribute('aria-expanded', open ? 'true' : 'false');
        });
        // ESC pour fermer
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape' && document.body.classList.contains('is-nav-open')) {
                document.body.classList.remove('is-nav-open');
                burger.setAttribute('aria-expanded', 'false');
            }
        });
    }

    /* ---- 3. Mobile menu accordion ---- */
    function setupMobileAccordion() {
        var items = document.querySelectorAll('.topnav-mobile-item');
        items.forEach(function (item) {
            var btn = item.querySelector('button.topnav-mobile-link');
            if (!btn) return; // pas de sub-items
            btn.addEventListener('click', function (e) {
                e.preventDefault();
                // Ferme les autres
                items.forEach(function (other) {
                    if (other !== item) other.classList.remove('is-open');
                });
                item.classList.toggle('is-open');
            });
        });
    }

    /* ---- 4. Desktop dropdowns avec grace period (tolère le transit trigger→panel) ---- */
    function setupDesktopDropdowns() {
        var items = document.querySelectorAll('.topnav-item.has-panel');
        if (!items.length) return;

        var GRACE = 180; // ms avant fermeture après mouseleave

        items.forEach(function (item) {
            var panel = item.querySelector('.topnav-panel');
            var closeT = null;

            function open() {
                clearTimeout(closeT);
                // Ferme les autres panels ouverts
                items.forEach(function (other) {
                    if (other !== item) other.classList.remove('is-hover');
                });
                item.classList.add('is-hover');
                document.body.classList.add('is-nav-open');
            }
            function scheduleClose() {
                clearTimeout(closeT);
                closeT = setTimeout(function () {
                    item.classList.remove('is-hover');
                    if (!document.querySelector('.topnav-item.is-hover')) {
                        // Si pas en mode burger ouvert, retire la classe body
                        if (!document.querySelector('.topbar-burger[aria-expanded="true"]')) {
                            document.body.classList.remove('is-nav-open');
                        }
                    }
                }, GRACE);
            }

            item.addEventListener('mouseenter', open);
            item.addEventListener('mouseleave', scheduleClose);
            if (panel) {
                panel.addEventListener('mouseenter', open);
                panel.addEventListener('mouseleave', scheduleClose);
            }
        });

        // Ferme tous les panels en cliquant ailleurs ou via ESC
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') {
                items.forEach(function (it) { it.classList.remove('is-hover'); });
                if (!document.querySelector('.topbar-burger[aria-expanded="true"]')) {
                    document.body.classList.remove('is-nav-open');
                }
            }
        });
    }

    /* ---- 5. Fermer le menu mobile en cliquant sur un lien ---- */
    function setupMobileLinkClose() {
        var links = document.querySelectorAll('.topnav-mobile-sub a, .topnav-mobile-foot a');
        links.forEach(function (link) {
            link.addEventListener('click', function () {
                document.body.classList.remove('is-nav-open');
                var burger = document.querySelector('.topbar-burger');
                if (burger) burger.setAttribute('aria-expanded', 'false');
            });
        });
    }

    /* ---- 6. Scroll-reveal (utilisé sur certaines pages) ---- */
    function setupReveals() {
        var els = document.querySelectorAll('[data-reveal], [data-reveal-photo]');
        if (!els.length || !('IntersectionObserver' in window)) return;
        var prefersReduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
        if (prefersReduced) {
            els.forEach(function (el) { el.classList.add('is-revealed'); });
            return;
        }
        var obs = new IntersectionObserver(function (entries) {
            entries.forEach(function (entry) {
                if (entry.isIntersecting) {
                    entry.target.classList.add('is-revealed');
                    obs.unobserve(entry.target);
                }
            });
        }, { rootMargin: '0px 0px -8% 0px', threshold: 0.05 });
        els.forEach(function (el) { obs.observe(el); });
    }

    /* ---- 7. Modale légale (Mentions / Confidentialité / CGV) ---- */
    function setupLegalModal() {
        var triggers = document.querySelectorAll('[data-modal-open]');
        var modal = document.getElementById('legal-modal');
        if (!modal) return;
        var closeButtons = modal.querySelectorAll('[data-modal-close]');
        var tabs = modal.querySelectorAll('.modal-tab');
        var panes = modal.querySelectorAll('.modal-pane');

        function open(targetTab) {
            document.body.classList.add('is-modal-open');
            modal.removeAttribute('hidden');
            modal.setAttribute('aria-hidden', 'false');
            if (targetTab) switchTo(targetTab);
            // Focus le close pour a11y
            setTimeout(function () {
                var c = modal.querySelector('.modal-close');
                if (c) c.focus();
            }, 100);
        }
        function close() {
            document.body.classList.remove('is-modal-open');
            modal.setAttribute('aria-hidden', 'true');
            // delai pour laisser l'anim de sortie (mais on a pas d'anim sortie ici, simple cut)
            setTimeout(function () {
                if (!document.body.classList.contains('is-modal-open')) {
                    modal.setAttribute('hidden', '');
                }
            }, 50);
        }
        function switchTo(name) {
            tabs.forEach(function (t) {
                t.classList.toggle('is-active', t.dataset.tab === name);
            });
            panes.forEach(function (p) {
                p.classList.toggle('is-active', p.dataset.pane === name);
            });
            // Scroll body to top
            var body = modal.querySelector('.modal-body');
            if (body) body.scrollTop = 0;
        }

        triggers.forEach(function (trigger) {
            trigger.addEventListener('click', function (e) {
                e.preventDefault();
                var target = trigger.dataset.modalOpen || 'mentions';
                open(target);
            });
        });
        closeButtons.forEach(function (btn) {
            btn.addEventListener('click', function (e) {
                e.preventDefault();
                close();
            });
        });
        tabs.forEach(function (tab) {
            tab.addEventListener('click', function (e) {
                e.preventDefault();
                switchTo(tab.dataset.tab);
            });
        });
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape' && document.body.classList.contains('is-modal-open')) {
                close();
            }
        });
    }

    /* ---- 8. Bandeau dev mode (toggle annotations photo) ---- */
    function setupDevBanner() {
        // Vérifie si la page contient des emplacements photo en attente
        var hasPending = document.querySelector('.photo-pending');
        if (!hasPending) return;

        // Crée le bandeau
        var banner = document.createElement('div');
        banner.className = 'dev-banner';
        banner.textContent = 'Aperçu placeholders';
        banner.title = 'Cliquer pour afficher / masquer les annotations photo';
        banner.addEventListener('click', function () {
            document.body.classList.toggle('no-dev-mode');
            try {
                localStorage.setItem('verrerie-no-dev-mode',
                    document.body.classList.contains('no-dev-mode') ? '1' : '0');
            } catch (e) { /* ignore */ }
        });
        document.body.appendChild(banner);

        // Restaure l'état précédent
        try {
            if (localStorage.getItem('verrerie-no-dev-mode') === '1') {
                document.body.classList.add('no-dev-mode');
            }
        } catch (e) { /* ignore */ }
    }

    /* ---- Init ---- */
    function init() {
        setupScrollState();
        setupBurger();
        setupMobileAccordion();
        setupDesktopDropdowns();
        setupMobileLinkClose();
        setupReveals();
        setupLegalModal();
        setupDevBanner();
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
