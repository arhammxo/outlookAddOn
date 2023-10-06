!function(){"use strict";var n,A,e,t,i,o,r,a,c,d,s,l,p,f,g,u,h,E,m,w,b,B,y,v={46556:function(n,A,e){var t=e(87537),i=e.n(t),o=e(23645),r=e.n(o)()(i());r.push([n.id,'/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. */\n\n/* ! tailwindcss v3.3.3 | MIT License | https://tailwindcss.com */\n\n/*\n1. Prevent padding and border from affecting element width. (https://github.com/mozdevs/cssremedy/issues/4)\n2. Allow adding a border to an element by just adding a border-width. (https://github.com/tailwindcss/tailwindcss/pull/116)\n*/\n\n*,\n::before,\n::after {\n  box-sizing: border-box; /* 1 */\n  border-width: 0; /* 2 */\n  border-style: solid; /* 2 */\n  border-color: #e5e7eb; /* 2 */\n}\n\n::before,\n::after {\n  --tw-content: \'\';\n}\n\n/*\n1. Use a consistent sensible line-height in all browsers.\n2. Prevent adjustments of font size after orientation changes in iOS.\n3. Use a more readable tab size.\n4. Use the user\'s configured `sans` font-family by default.\n5. Use the user\'s configured `sans` font-feature-settings by default.\n6. Use the user\'s configured `sans` font-variation-settings by default.\n*/\n\nhtml {\n  line-height: 1.5; /* 1 */\n  -webkit-text-size-adjust: 100%; /* 2 */ /* 3 */\n  tab-size: 4; /* 3 */\n  font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Ubuntu, Cantarell, Noto Sans, sans-serif, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans", sans-serif, "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol", "Noto Color Emoji"; /* 4 */\n  font-feature-settings: normal; /* 5 */\n  font-variation-settings: normal; /* 6 */\n}\n\n/*\n1. Remove the margin in all browsers.\n2. Inherit line-height from `html` so users can set them as a class directly on the `html` element.\n*/\n\nbody {\n  margin: 0; /* 1 */\n  line-height: inherit; /* 2 */\n}\n\n/*\n1. Add the correct height in Firefox.\n2. Correct the inheritance of border color in Firefox. (https://bugzilla.mozilla.org/show_bug.cgi?id=190655)\n3. Ensure horizontal rules are visible by default.\n*/\n\nhr {\n  height: 0; /* 1 */\n  color: inherit; /* 2 */\n  border-top-width: 1px; /* 3 */\n}\n\n/*\nAdd the correct text decoration in Chrome, Edge, and Safari.\n*/\n\nabbr:where([title]) {\n  text-decoration: underline;\n  text-decoration: underline dotted;\n}\n\n/*\nRemove the default font size and weight for headings.\n*/\n\nh1,\nh2,\nh3,\nh4,\nh5,\nh6 {\n  font-size: inherit;\n  font-weight: inherit;\n}\n\n/*\nReset links to optimize for opt-in styling instead of opt-out.\n*/\n\na {\n  color: inherit;\n  text-decoration: inherit;\n}\n\n/*\nAdd the correct font weight in Edge and Safari.\n*/\n\nb,\nstrong {\n  font-weight: bolder;\n}\n\n/*\n1. Use the user\'s configured `mono` font family by default.\n2. Correct the odd `em` font sizing in all browsers.\n*/\n\ncode,\nkbd,\nsamp,\npre {\n  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace; /* 1 */\n  font-size: 1em; /* 2 */\n}\n\n/*\nAdd the correct font size in all browsers.\n*/\n\nsmall {\n  font-size: 80%;\n}\n\n/*\nPrevent `sub` and `sup` elements from affecting the line height in all browsers.\n*/\n\nsub,\nsup {\n  font-size: 75%;\n  line-height: 0;\n  position: relative;\n  vertical-align: baseline;\n}\n\nsub {\n  bottom: -0.25em;\n}\n\nsup {\n  top: -0.5em;\n}\n\n/*\n1. Remove text indentation from table contents in Chrome and Safari. (https://bugs.chromium.org/p/chromium/issues/detail?id=999088, https://bugs.webkit.org/show_bug.cgi?id=201297)\n2. Correct table border color inheritance in all Chrome and Safari. (https://bugs.chromium.org/p/chromium/issues/detail?id=935729, https://bugs.webkit.org/show_bug.cgi?id=195016)\n3. Remove gaps between table borders by default.\n*/\n\ntable {\n  text-indent: 0; /* 1 */\n  border-color: inherit; /* 2 */\n  border-collapse: collapse; /* 3 */\n}\n\n/*\n1. Change the font styles in all browsers.\n2. Remove the margin in Firefox and Safari.\n3. Remove default padding in all browsers.\n*/\n\nbutton,\ninput,\noptgroup,\nselect,\ntextarea {\n  font-family: inherit; /* 1 */\n  font-feature-settings: inherit; /* 1 */\n  font-variation-settings: inherit; /* 1 */\n  font-size: 100%; /* 1 */\n  font-weight: inherit; /* 1 */\n  line-height: inherit; /* 1 */\n  color: inherit; /* 1 */\n  margin: 0; /* 2 */\n  padding: 0; /* 3 */\n}\n\n/*\nRemove the inheritance of text transform in Edge and Firefox.\n*/\n\nbutton,\nselect {\n  text-transform: none;\n}\n\n/*\n1. Correct the inability to style clickable types in iOS and Safari.\n2. Remove default button styles.\n*/\n\nbutton,\n[type=\'button\'],\n[type=\'reset\'],\n[type=\'submit\'] {\n  -webkit-appearance: button; /* 1 */\n  background-color: transparent; /* 2 */\n  background-image: none; /* 2 */\n}\n\n/*\nUse the modern Firefox focus style for all focusable elements.\n*/\n\n:-moz-focusring {\n  outline: auto;\n}\n\n/*\nRemove the additional `:invalid` styles in Firefox. (https://github.com/mozilla/gecko-dev/blob/2f9eacd9d3d995c937b4251a5557d95d494c9be1/layout/style/res/forms.css#L728-L737)\n*/\n\n:-moz-ui-invalid {\n  box-shadow: none;\n}\n\n/*\nAdd the correct vertical alignment in Chrome and Firefox.\n*/\n\nprogress {\n  vertical-align: baseline;\n}\n\n/*\nCorrect the cursor style of increment and decrement buttons in Safari.\n*/\n\n::-webkit-inner-spin-button,\n::-webkit-outer-spin-button {\n  height: auto;\n}\n\n/*\n1. Correct the odd appearance in Chrome and Safari.\n2. Correct the outline style in Safari.\n*/\n\n[type=\'search\'] {\n  -webkit-appearance: textfield; /* 1 */\n  outline-offset: -2px; /* 2 */\n}\n\n/*\nRemove the inner padding in Chrome and Safari on macOS.\n*/\n\n::-webkit-search-decoration {\n  -webkit-appearance: none;\n}\n\n/*\n1. Correct the inability to style clickable types in iOS and Safari.\n2. Change font properties to `inherit` in Safari.\n*/\n\n::-webkit-file-upload-button {\n  -webkit-appearance: button; /* 1 */\n  font: inherit; /* 2 */\n}\n\n/*\nAdd the correct display in Chrome and Safari.\n*/\n\nsummary {\n  display: list-item;\n}\n\n/*\nRemoves the default spacing and border for appropriate elements.\n*/\n\nblockquote,\ndl,\ndd,\nh1,\nh2,\nh3,\nh4,\nh5,\nh6,\nhr,\nfigure,\np,\npre {\n  margin: 0;\n}\n\nfieldset {\n  margin: 0;\n  padding: 0;\n}\n\nlegend {\n  padding: 0;\n}\n\nol,\nul,\nmenu {\n  list-style: none;\n  margin: 0;\n  padding: 0;\n}\n\n/*\nReset default styling for dialogs.\n*/\n\ndialog {\n  padding: 0;\n}\n\n/*\nPrevent resizing textareas horizontally by default.\n*/\n\ntextarea {\n  resize: vertical;\n}\n\n/*\n1. Reset the default placeholder opacity in Firefox. (https://github.com/tailwindlabs/tailwindcss/issues/3300)\n2. Set the default placeholder color to the user\'s configured gray 400 color.\n*/\n\ninput:-ms-input-placeholder, textarea:-ms-input-placeholder {\n  opacity: 1; /* 1 */\n  color: #9ca3af; /* 2 */\n}\n\ninput::placeholder,\ntextarea::placeholder {\n  opacity: 1; /* 1 */\n  color: #9ca3af; /* 2 */\n}\n\n/*\nSet the default cursor for buttons.\n*/\n\nbutton,\n[role="button"] {\n  cursor: pointer;\n}\n\n/*\nMake sure disabled buttons don\'t get the pointer cursor.\n*/\n\n:disabled {\n  cursor: default;\n}\n\n/*\n1. Make replaced elements `display: block` by default. (https://github.com/mozdevs/cssremedy/issues/14)\n2. Add `vertical-align: middle` to align replaced elements more sensibly by default. (https://github.com/jensimmons/cssremedy/issues/14#issuecomment-634934210)\n   This can trigger a poorly considered lint error in some tools but is included by design.\n*/\n\nimg,\nsvg,\nvideo,\ncanvas,\naudio,\niframe,\nembed,\nobject {\n  display: block; /* 1 */\n  vertical-align: middle; /* 2 */\n}\n\n/*\nConstrain images and videos to the parent width and preserve their intrinsic aspect ratio. (https://github.com/mozdevs/cssremedy/issues/14)\n*/\n\nimg,\nvideo {\n  max-width: 100%;\n  height: auto;\n}\n\n/* Make elements with the HTML hidden attribute stay hidden by default */\n\n[hidden] {\n  display: none;\n}\n\n*, ::before, ::after {\n  --tw-border-spacing-x: 0;\n  --tw-border-spacing-y: 0;\n  --tw-translate-x: 0;\n  --tw-translate-y: 0;\n  --tw-rotate: 0;\n  --tw-skew-x: 0;\n  --tw-skew-y: 0;\n  --tw-scale-x: 1;\n  --tw-scale-y: 1;\n  --tw-pan-x:  ;\n  --tw-pan-y:  ;\n  --tw-pinch-zoom:  ;\n  --tw-scroll-snap-strictness: proximity;\n  --tw-gradient-from-position:  ;\n  --tw-gradient-via-position:  ;\n  --tw-gradient-to-position:  ;\n  --tw-ordinal:  ;\n  --tw-slashed-zero:  ;\n  --tw-numeric-figure:  ;\n  --tw-numeric-spacing:  ;\n  --tw-numeric-fraction:  ;\n  --tw-ring-inset:  ;\n  --tw-ring-offset-width: 0px;\n  --tw-ring-offset-color: #fff;\n  --tw-ring-color: rgba(59, 130, 246, 0.5);\n  --tw-ring-offset-shadow: 0 0 rgba(0,0,0,0);\n  --tw-ring-shadow: 0 0 rgba(0,0,0,0);\n  --tw-shadow: 0 0 rgba(0,0,0,0);\n  --tw-shadow-colored: 0 0 rgba(0,0,0,0);\n  --tw-blur:  ;\n  --tw-brightness:  ;\n  --tw-contrast:  ;\n  --tw-grayscale:  ;\n  --tw-hue-rotate:  ;\n  --tw-invert:  ;\n  --tw-saturate:  ;\n  --tw-sepia:  ;\n  --tw-drop-shadow:  ;\n  --tw-backdrop-blur:  ;\n  --tw-backdrop-brightness:  ;\n  --tw-backdrop-contrast:  ;\n  --tw-backdrop-grayscale:  ;\n  --tw-backdrop-hue-rotate:  ;\n  --tw-backdrop-invert:  ;\n  --tw-backdrop-opacity:  ;\n  --tw-backdrop-saturate:  ;\n  --tw-backdrop-sepia:  ;\n}\n\n::-ms-backdrop {\n  --tw-border-spacing-x: 0;\n  --tw-border-spacing-y: 0;\n  --tw-translate-x: 0;\n  --tw-translate-y: 0;\n  --tw-rotate: 0;\n  --tw-skew-x: 0;\n  --tw-skew-y: 0;\n  --tw-scale-x: 1;\n  --tw-scale-y: 1;\n  --tw-pan-x:  ;\n  --tw-pan-y:  ;\n  --tw-pinch-zoom:  ;\n  --tw-scroll-snap-strictness: proximity;\n  --tw-gradient-from-position:  ;\n  --tw-gradient-via-position:  ;\n  --tw-gradient-to-position:  ;\n  --tw-ordinal:  ;\n  --tw-slashed-zero:  ;\n  --tw-numeric-figure:  ;\n  --tw-numeric-spacing:  ;\n  --tw-numeric-fraction:  ;\n  --tw-ring-inset:  ;\n  --tw-ring-offset-width: 0px;\n  --tw-ring-offset-color: #fff;\n  --tw-ring-color: rgba(59, 130, 246, 0.5);\n  --tw-ring-offset-shadow: 0 0 rgba(0,0,0,0);\n  --tw-ring-shadow: 0 0 rgba(0,0,0,0);\n  --tw-shadow: 0 0 rgba(0,0,0,0);\n  --tw-shadow-colored: 0 0 rgba(0,0,0,0);\n  --tw-blur:  ;\n  --tw-brightness:  ;\n  --tw-contrast:  ;\n  --tw-grayscale:  ;\n  --tw-hue-rotate:  ;\n  --tw-invert:  ;\n  --tw-saturate:  ;\n  --tw-sepia:  ;\n  --tw-drop-shadow:  ;\n  --tw-backdrop-blur:  ;\n  --tw-backdrop-brightness:  ;\n  --tw-backdrop-contrast:  ;\n  --tw-backdrop-grayscale:  ;\n  --tw-backdrop-hue-rotate:  ;\n  --tw-backdrop-invert:  ;\n  --tw-backdrop-opacity:  ;\n  --tw-backdrop-saturate:  ;\n  --tw-backdrop-sepia:  ;\n}\n\n::backdrop {\n  --tw-border-spacing-x: 0;\n  --tw-border-spacing-y: 0;\n  --tw-translate-x: 0;\n  --tw-translate-y: 0;\n  --tw-rotate: 0;\n  --tw-skew-x: 0;\n  --tw-skew-y: 0;\n  --tw-scale-x: 1;\n  --tw-scale-y: 1;\n  --tw-pan-x:  ;\n  --tw-pan-y:  ;\n  --tw-pinch-zoom:  ;\n  --tw-scroll-snap-strictness: proximity;\n  --tw-gradient-from-position:  ;\n  --tw-gradient-via-position:  ;\n  --tw-gradient-to-position:  ;\n  --tw-ordinal:  ;\n  --tw-slashed-zero:  ;\n  --tw-numeric-figure:  ;\n  --tw-numeric-spacing:  ;\n  --tw-numeric-fraction:  ;\n  --tw-ring-inset:  ;\n  --tw-ring-offset-width: 0px;\n  --tw-ring-offset-color: #fff;\n  --tw-ring-color: rgba(59, 130, 246, 0.5);\n  --tw-ring-offset-shadow: 0 0 rgba(0,0,0,0);\n  --tw-ring-shadow: 0 0 rgba(0,0,0,0);\n  --tw-shadow: 0 0 rgba(0,0,0,0);\n  --tw-shadow-colored: 0 0 rgba(0,0,0,0);\n  --tw-blur:  ;\n  --tw-brightness:  ;\n  --tw-contrast:  ;\n  --tw-grayscale:  ;\n  --tw-hue-rotate:  ;\n  --tw-invert:  ;\n  --tw-saturate:  ;\n  --tw-sepia:  ;\n  --tw-drop-shadow:  ;\n  --tw-backdrop-blur:  ;\n  --tw-backdrop-brightness:  ;\n  --tw-backdrop-contrast:  ;\n  --tw-backdrop-grayscale:  ;\n  --tw-backdrop-hue-rotate:  ;\n  --tw-backdrop-invert:  ;\n  --tw-backdrop-opacity:  ;\n  --tw-backdrop-saturate:  ;\n  --tw-backdrop-sepia:  ;\n}\n.block {\n  display: block;\n}\n.contents {\n  display: contents;\n}\n.font-bold {\n  font-weight: 700;\n}\n.filter {\n  filter: var(--tw-blur) var(--tw-brightness) var(--tw-contrast) var(--tw-grayscale) var(--tw-hue-rotate) var(--tw-invert) var(--tw-saturate) var(--tw-sepia) var(--tw-drop-shadow);\n}\nhtml,\nbody {\n    width: 100%;\n    height: 100%;\n    margin: 0;\n    padding: 0;\n    overflow: auto;\n    position: relative;\n    font-size: 16px;\n}\n\nmain {\n    height: 100%;\n    overflow-y: auto;\n    background: white;\n}\n\nfooter {\n    width: 100%;\n    position: relative;\n    bottom: 0;\n    margin-top: 10px;\n}\n\np,\nh1,\nh2,\nh3,\nh4,\nh5,\nh6 {\n    margin: 0;\n    padding: 0;\n}\n\nul {\n    padding: 0;\n}\n\n#settings-prompt {\n    margin: 10px 0;\n}\n\n#error-display {\n    padding: 10px;\n}\n\n#insert-button {\n    margin: 0 10px;\n}\n\n.clearfix {\n    display: block;\n    clear: both;\n    height: 0;\n}\n\n.pointerCursor {\n    cursor: pointer;\n}\n\n.invisible {\n    visibility: hidden;\n}\n\n.undisplayed {\n    display: none;\n}\n\n.ms-Icon.enlarge {\n    position: relative;\n    font-size: 20px;\n    top: 4px;\n}\n\n.ms-ListItem-secondaryText,\n.ms-ListItem-tertiaryText {\n    padding-left: 15px;\n}\n\n.ms-landing-page {\n    display: flex;\n    flex-direction: column;\n    flex-wrap: nowrap;\n    height: 100%;\n}\n\n.ms-landing-page__main {\n    display: flex;\n    flex-direction: column;\n    flex-wrap: nowrap;\n    flex: 1 1 0;\n    height: 100%;\n}\n\n.ms-landing-page__content {\n    display: flex;\n    flex-direction: column;\n    flex-wrap: nowrap;\n    height: 100%;\n    flex: 1 1 0;\n    padding: 20px;\n}\n\n.ms-landing-page__content h2 {\n    margin-bottom: 20px;\n}\n\n.ms-landing-page__footer {\n    display: inline-flex;\n    justify-content: center;\n    align-items: center;\n}\n\n.ms-landing-page__footer--left {\n    transition: background ease 0.1s, color ease 0.1s;\n    display: inline-flex;\n    justify-content: flex-start;\n    align-items: center;\n    flex: 1 0 0px;\n    padding: 20px;\n}\n\n.ms-landing-page__footer--left:active {\n    cursor: default;\n}\n\n.ms-landing-page__footer--left--disabled {\n    opacity: 0.6;\n    pointer-events: none;\n    cursor: not-allowed;\n}\n\n.ms-landing-page__footer--left--disabled:active,\n.ms-landing-page__footer--left--disabled:hover {\n    background: transparent;\n}\n\n.ms-landing-page__footer--left img {\n    width: 40px;\n    height: 40px;\n}\n\n.ms-landing-page__footer--left h1 {\n    flex: 1 0 0px;\n    margin-left: 15px;\n    text-align: left;\n    width: auto;\n    max-width: auto;\n    overflow: hidden;\n    white-space: nowrap;\n    text-overflow: ellipsis;\n}\n\n.ms-landing-page__footer--right {\n    transition: background ease 0.1s, color ease 0.1s;\n    padding: 29px 20px;\n}\n\n.ms-landing-page__footer--right:active,\n.ms-landing-page__footer--right:hover {\n    background: #005ca4;\n    cursor: pointer;\n}\n\n.ms-landing-page__footer--right:active {\n    background: #005ca4;\n}\n\n.ms-landing-page__footer--right--disabled {\n    opacity: 0.6;\n    pointer-events: none;\n    cursor: not-allowed;\n}\n\n.ms-landing-page__footer--right--disabled:active,\n.ms-landing-page__footer--right--disabled:hover {\n    background: transparent;\n}',"",{version:3,sources:["webpack://./src/taskpane/taskpane.css"],names:[],mappings:"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",sourcesContent:["/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. */\n\n@tailwind base;\n@tailwind components;\n@tailwind utilities;\nhtml,\nbody {\n    width: 100%;\n    height: 100%;\n    margin: 0;\n    padding: 0;\n    overflow: auto;\n    position: relative;\n    font-size: 16px;\n}\n\nmain {\n    height: 100%;\n    overflow-y: auto;\n    background: white;\n}\n\nfooter {\n    width: 100%;\n    position: relative;\n    bottom: 0;\n    margin-top: 10px;\n}\n\np,\nh1,\nh2,\nh3,\nh4,\nh5,\nh6 {\n    margin: 0;\n    padding: 0;\n}\n\nul {\n    padding: 0;\n}\n\n#settings-prompt {\n    margin: 10px 0;\n}\n\n#error-display {\n    padding: 10px;\n}\n\n#insert-button {\n    margin: 0 10px;\n}\n\n.clearfix {\n    display: block;\n    clear: both;\n    height: 0;\n}\n\n.pointerCursor {\n    cursor: pointer;\n}\n\n.invisible {\n    visibility: hidden;\n}\n\n.undisplayed {\n    display: none;\n}\n\n.ms-Icon.enlarge {\n    position: relative;\n    font-size: 20px;\n    top: 4px;\n}\n\n.ms-ListItem-secondaryText,\n.ms-ListItem-tertiaryText {\n    padding-left: 15px;\n}\n\n.ms-landing-page {\n    display: -webkit-flex;\n    display: flex;\n    -webkit-flex-direction: column;\n    flex-direction: column;\n    -webkit-flex-wrap: nowrap;\n    flex-wrap: nowrap;\n    height: 100%;\n}\n\n.ms-landing-page__main {\n    display: -webkit-flex;\n    display: flex;\n    -webkit-flex-direction: column;\n    flex-direction: column;\n    -webkit-flex-wrap: nowrap;\n    flex-wrap: nowrap;\n    -webkit-flex: 1 1 0;\n    flex: 1 1 0;\n    height: 100%;\n}\n\n.ms-landing-page__content {\n    display: -webkit-flex;\n    display: flex;\n    -webkit-flex-direction: column;\n    flex-direction: column;\n    -webkit-flex-wrap: nowrap;\n    flex-wrap: nowrap;\n    height: 100%;\n    -webkit-flex: 1 1 0;\n    flex: 1 1 0;\n    padding: 20px;\n}\n\n.ms-landing-page__content h2 {\n    margin-bottom: 20px;\n}\n\n.ms-landing-page__footer {\n    display: -webkit-inline-flex;\n    display: inline-flex;\n    -webkit-justify-content: center;\n    justify-content: center;\n    -webkit-align-items: center;\n    align-items: center;\n}\n\n.ms-landing-page__footer--left {\n    transition: background ease 0.1s, color ease 0.1s;\n    display: -webkit-inline-flex;\n    display: inline-flex;\n    -webkit-justify-content: flex-start;\n    justify-content: flex-start;\n    -webkit-align-items: center;\n    align-items: center;\n    -webkit-flex: 1 0 0px;\n    flex: 1 0 0px;\n    padding: 20px;\n}\n\n.ms-landing-page__footer--left:active {\n    cursor: default;\n}\n\n.ms-landing-page__footer--left--disabled {\n    opacity: 0.6;\n    pointer-events: none;\n    cursor: not-allowed;\n}\n\n.ms-landing-page__footer--left--disabled:active,\n.ms-landing-page__footer--left--disabled:hover {\n    background: transparent;\n}\n\n.ms-landing-page__footer--left img {\n    width: 40px;\n    height: 40px;\n}\n\n.ms-landing-page__footer--left h1 {\n    -webkit-flex: 1 0 0px;\n    flex: 1 0 0px;\n    margin-left: 15px;\n    text-align: left;\n    width: auto;\n    max-width: auto;\n    overflow: hidden;\n    white-space: nowrap;\n    text-overflow: ellipsis;\n}\n\n.ms-landing-page__footer--right {\n    transition: background ease 0.1s, color ease 0.1s;\n    padding: 29px 20px;\n}\n\n.ms-landing-page__footer--right:active,\n.ms-landing-page__footer--right:hover {\n    background: #005ca4;\n    cursor: pointer;\n}\n\n.ms-landing-page__footer--right:active {\n    background: #005ca4;\n}\n\n.ms-landing-page__footer--right--disabled {\n    opacity: 0.6;\n    pointer-events: none;\n    cursor: not-allowed;\n}\n\n.ms-landing-page__footer--right--disabled:active,\n.ms-landing-page__footer--right--disabled:hover {\n    background: transparent;\n}"],sourceRoot:""}]),A.Z=r},23645:function(n){n.exports=function(n){var A=[];return A.toString=function(){return this.map((function(A){var e="",t=void 0!==A[5];return A[4]&&(e+="@supports (".concat(A[4],") {")),A[2]&&(e+="@media ".concat(A[2]," {")),t&&(e+="@layer".concat(A[5].length>0?" ".concat(A[5]):""," {")),e+=n(A),t&&(e+="}"),A[2]&&(e+="}"),A[4]&&(e+="}"),e})).join("")},A.i=function(n,e,t,i,o){"string"==typeof n&&(n=[[null,n,void 0]]);var r={};if(t)for(var a=0;a<this.length;a++){var c=this[a][0];null!=c&&(r[c]=!0)}for(var d=0;d<n.length;d++){var s=[].concat(n[d]);t&&r[s[0]]||(void 0!==o&&(void 0===s[5]||(s[1]="@layer".concat(s[5].length>0?" ".concat(s[5]):""," {").concat(s[1],"}")),s[5]=o),e&&(s[2]?(s[1]="@media ".concat(s[2]," {").concat(s[1],"}"),s[2]=e):s[2]=e),i&&(s[4]?(s[1]="@supports (".concat(s[4],") {").concat(s[1],"}"),s[4]=i):s[4]="".concat(i)),A.push(s))}},A}},87537:function(n){n.exports=function(n){var A=n[1],e=n[3];if(!e)return A;if("function"==typeof btoa){var t=btoa(unescape(encodeURIComponent(JSON.stringify(e)))),i="sourceMappingURL=data:application/json;charset=utf-8;base64,".concat(t),o="/*# ".concat(i," */");return[A].concat([o]).join("\n")}return[A].join("\n")}},27091:function(n){n.exports=function(n,A){return A||(A={}),n?(n=String(n.__esModule?n.default:n),A.hash&&(n+=A.hash),A.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(n)?'"'.concat(n,'"'):n):n}},93379:function(n){var A=[];function e(n){for(var e=-1,t=0;t<A.length;t++)if(A[t].identifier===n){e=t;break}return e}function t(n,t){for(var o={},r=[],a=0;a<n.length;a++){var c=n[a],d=t.base?c[0]+t.base:c[0],s=o[d]||0,l="".concat(d," ").concat(s);o[d]=s+1;var p=e(l),f={css:c[1],media:c[2],sourceMap:c[3],supports:c[4],layer:c[5]};if(-1!==p)A[p].references++,A[p].updater(f);else{var g=i(f,t);t.byIndex=a,A.splice(a,0,{identifier:l,updater:g,references:1})}r.push(l)}return r}function i(n,A){var e=A.domAPI(A);return e.update(n),function(A){if(A){if(A.css===n.css&&A.media===n.media&&A.sourceMap===n.sourceMap&&A.supports===n.supports&&A.layer===n.layer)return;e.update(n=A)}else e.remove()}}n.exports=function(n,i){var o=t(n=n||[],i=i||{});return function(n){n=n||[];for(var r=0;r<o.length;r++){var a=e(o[r]);A[a].references--}for(var c=t(n,i),d=0;d<o.length;d++){var s=e(o[d]);0===A[s].references&&(A[s].updater(),A.splice(s,1))}o=c}}},90569:function(n){var A={};n.exports=function(n,e){var t=function(n){if(void 0===A[n]){var e=document.querySelector(n);if(window.HTMLIFrameElement&&e instanceof window.HTMLIFrameElement)try{e=e.contentDocument.head}catch(n){e=null}A[n]=e}return A[n]}(n);if(!t)throw new Error("Couldn't find a style target. This probably means that the value for the 'insert' parameter is invalid.");t.appendChild(e)}},19216:function(n){n.exports=function(n){var A=document.createElement("style");return n.setAttributes(A,n.attributes),n.insert(A,n.options),A}},3565:function(n,A,e){n.exports=function(n){var A=e.nc;A&&n.setAttribute("nonce",A)}},7795:function(n){n.exports=function(n){if("undefined"==typeof document)return{update:function(){},remove:function(){}};var A=n.insertStyleElement(n);return{update:function(e){!function(n,A,e){var t="";e.supports&&(t+="@supports (".concat(e.supports,") {")),e.media&&(t+="@media ".concat(e.media," {"));var i=void 0!==e.layer;i&&(t+="@layer".concat(e.layer.length>0?" ".concat(e.layer):""," {")),t+=e.css,i&&(t+="}"),e.media&&(t+="}"),e.supports&&(t+="}");var o=e.sourceMap;o&&"undefined"!=typeof btoa&&(t+="\n/*# sourceMappingURL=data:application/json;base64,".concat(btoa(unescape(encodeURIComponent(JSON.stringify(o))))," */")),A.styleTagTransform(t,n,A.options)}(A,n,e)},remove:function(){!function(n){if(null===n.parentNode)return!1;n.parentNode.removeChild(n)}(A)}}}},44589:function(n){n.exports=function(n,A){if(A.styleSheet)A.styleSheet.cssText=n;else{for(;A.firstChild;)A.removeChild(A.firstChild);A.appendChild(document.createTextNode(n))}}},44944:function(n,A,e){n.exports=e.p+"assets/logo-filled.png"},49499:function(n,A,e){n.exports=e.p+"85b83baccb34ce0cf6b4.js"},41149:function(n,A,e){n.exports=e.p+"2321f24e69e0eac80628.js"},36076:function(n,A,e){n.exports=e.p+"a031ee3cf9ec85c43b50.js"},58038:function(n,A,e){n.exports=e.p+"4c536e224b0b3348659b.js"},79254:function(n,A,e){n.exports=e.p+"fb4b72b38f72ea6fb909.js"},66039:function(n,A,e){n.exports=e.p+"4cc5d58837ef495cacb2.js"}},C={};function x(n){var A=C[n];if(void 0!==A)return A.exports;var e=C[n]={id:n,exports:{}};return v[n](e,e.exports,x),e.exports}x.m=v,x.n=function(n){var A=n&&n.__esModule?function(){return n.default}:function(){return n};return x.d(A,{a:A}),A},x.d=function(n,A){for(var e in A)x.o(A,e)&&!x.o(n,e)&&Object.defineProperty(n,e,{enumerable:!0,get:A[e]})},x.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(n){if("object"==typeof window)return window}}(),x.o=function(n,A){return Object.prototype.hasOwnProperty.call(n,A)},function(){var n;x.g.importScripts&&(n=x.g.location+"");var A=x.g.document;if(!n&&A&&(A.currentScript&&(n=A.currentScript.src),!n)){var e=A.getElementsByTagName("script");if(e.length)for(var t=e.length-1;t>-1&&!n;)n=e[t--].src}if(!n)throw new Error("Automatic publicPath is not supported in this browser");n=n.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),x.p=n}(),x.b=document.baseURI||self.location.href,x.nc=void 0,d=x(93379),s=x.n(d),l=x(7795),p=x.n(l),f=x(90569),g=x.n(f),u=x(3565),h=x.n(u),E=x(19216),m=x.n(E),w=x(44589),b=x.n(w),B=x(46556),(y={}).styleTagTransform=b(),y.setAttributes=h(),y.insert=g().bind(null,"head"),y.domAPI=p(),y.insertStyleElement=m(),s()(B.Z,y),B.Z&&B.Z.locals&&B.Z.locals,function(){var n,A;function e(n){$("#error-display").hide(),$("#not-configured").hide(),$("#gist-list-container").show(),getUserGists(n,(function(n,A){A||($("#gist-list").empty(),buildGistList($("#gist-list"),n,t))}))}function t(){$("#insert-button").removeAttr("disabled"),$(".ms-ListItem").removeClass("is-selected").removeAttr("checked"),$(this).children(".ms-ListItem").addClass("is-selected").attr("checked","checked")}function i(t){n=JSON.parse(t.message),setConfig(n,(function(t){A.close(),A=null,e(n.gitHubUserName)}))}function o(n){A=null}Office.initialize=function(t){jQuery(document).ready((function(){(n=getConfig())&&n.gitHubUserName?e(n.gitHubUserName):$("#not-configured").show(),$("#insert-button").on("click",(function(){var n=$(".ms-ListItem.is-selected").val();getGist(n,(function(n,A){Office.context.mailbox.item.body.setSelectedDataAsync(n,{coercionType:Office.CoercionType.Html},(function(n){n.status===Office.AsyncResultStatus.Failed&&function(n){$("#not-configured").hide(),$("#gist-list-container").hide(),$("#error-display").text(n),$("#error-display").show()}("Could not insert gist: "+n.error.message)}))}))})),$("#settings-icon").on("click",(function(){var e=new URI("dialog.html").absoluteTo(window.location).toString();n&&(e=e+"?gitHubUserName="+n.gitHubUserName+"&defaultGistId="+n.defaultGistId),Office.context.ui.displayDialogAsync(e,{width:20,height:40,displayInIframe:!0},(function(n){(A=n.value).addEventHandler(Office.EventType.DialogMessageReceived,i),A.addEventHandler(Office.EventType.DialogEventReceived,o)}))}))}))}}(),n=x(27091),A=x.n(n),e=new URL(x(44944),x.b),t=new URL(x(58038),x.b),i=new URL(x(79254),x.b),o=new URL(x(66039),x.b),r=new URL(x(49499),x.b),a=new URL(x(41149),x.b),c=new URL(x(36076),x.b),A()(e),A()(t),A()(i),A()(o),A()(r),A()(a),A()(c)}();
//# sourceMappingURL=taskpane.js.map