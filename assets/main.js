/* 交互：移动端导航、锚点高亮、作品分类筛选、滚动渐显 */

document.addEventListener('DOMContentLoaded', function () {

  /* 移动端导航开合（仅详情页有 .nav，首页无此元素会自动跳过） */
  var toggle = document.querySelector('.nav__toggle');
  var links = document.querySelector('.nav__links');
  if (toggle && links) {
    toggle.addEventListener('click', function () {
      links.classList.toggle('open');
    });
    links.addEventListener('click', function (e) {
      if (e.target.tagName === 'A') links.classList.remove('open');
    });
  }

  /* 滚动渐显 */
  var faders = document.querySelectorAll('.fade-in');
  if ('IntersectionObserver' in window) {
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (en.isIntersecting) { en.target.classList.add('visible'); io.unobserve(en.target); }
      });
    }, { threshold: 0.08, rootMargin: '0px 0px -40px 0px' });
    faders.forEach(function (el) { io.observe(el); });
  } else {
    faders.forEach(function (el) { el.classList.add('visible'); });
  }

  /* 当前区块高亮（左侧身份栏 sidenav） */
  var sections = [].slice.call(document.querySelectorAll('main section[id], main footer[id]'));
  var navMap = {};
  document.querySelectorAll('.sidenav a[href^="#"]').forEach(function (a) {
    navMap[a.getAttribute('href').slice(1)] = a;
  });
  if (sections.length && Object.keys(navMap).length && 'IntersectionObserver' in window) {
    var spy = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        var a = navMap[en.target.id];
        if (!a) return;
        if (en.isIntersecting) {
          Object.keys(navMap).forEach(function (k) { navMap[k].classList.remove('active'); });
          a.classList.add('active');
        }
      });
    }, { threshold: 0, rootMargin: '-45% 0px -50% 0px' });
    sections.forEach(function (s) { spy.observe(s); });
  }

  /* 通用分类筛选：.filter-bar[data-filter-target] 指向列表容器，data-row 指定行选择器
     作品列表行用 data-cat，游戏经历行用 data-pcat，按钮 data-filter 与之一致即可 */
  document.querySelectorAll('.filter-bar[data-filter-target]').forEach(function (bar) {
    var target = document.getElementById(bar.getAttribute('data-filter-target'));
    if (!target) return;
    var rowSel = bar.getAttribute('data-row') || '.projrow';
    var rows = [].slice.call(target.querySelectorAll(rowSel));
    var buttons = bar.querySelectorAll('.filter-btn');
    buttons.forEach(function (btn) {
      btn.addEventListener('click', function () {
        var f = btn.getAttribute('data-filter');
        buttons.forEach(function (b) { b.classList.remove('active'); });
        btn.classList.add('active');
        rows.forEach(function (r) {
          var show = (f === 'all') ||
            (r.getAttribute('data-cat') === f) ||
            (r.getAttribute('data-pcat') === f);
          r.style.display = show ? '' : 'none';
        });
      });
    });
  });

  /* 详情页自动目录 */
  var doc = document.querySelector('.doc');
  if (doc) {
    var navOl = document.querySelector('#toc-list');
    var heads = doc.querySelectorAll('h2');
    if (navOl && heads.length) {
      var i = 0;
      heads.forEach(function (h) {
        i += 1;
        var id = 'sec-' + i;
        h.setAttribute('id', id);
        var li = document.createElement('li');
        var a = document.createElement('a');
        a.href = '#' + id;
        a.textContent = h.textContent.trim();
        li.appendChild(a);
        navOl.appendChild(li);
      });
      var box = document.getElementById('toc-box');
      if (box) box.style.display = '';
    }
  }

});
