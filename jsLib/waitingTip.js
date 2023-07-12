/**
 *author: Hal
 */

/**
 * @praam <String>containerElId=" ____waiting____随机数" 指定一个容器的id
 * @param <String>styleClassName 容器的css样式类
 * @param <String>innerHTML ＝"<img alt='running...' src='/images/waiting.gif' /> "  内容
 * @param <String>anchor ="down"  停放位置 ["down","up","left","right","center"];
 * @param <Number>gap =2与参照节点位置的间距
 */
function WaitingTip(options) {
    if (!options) {
        options = {
            containerElId: null,
            styleClassName: null,
            innerHTML: null,
            anchor: null,
            gap: null
        };
    }
    var id = options.containerElid || " ____waiting" + Math.floor(Math.random() * 1000000);
    this.getWaitEl = function () {
        return document.getElementById(id);
    };

    var anchor = options.anchor ? options.anchor.toLowerCase() : "down";
    this.getAnchor = function () {
        return anchor;
    };

    var gap = options.gap || 2;
    this.getGap = function () {
        return gap;
    };

    var init = function () {
        var div = document.createElement("div");
        div.id = id;
        div.style.position = "absolute";
        div.style.display = "none";
        if (options.styleClassName) div.className = styleClassName;
        document.body.appendChild(div);
        if (options.innerHTML) {
            div.innerHTML = options.innerHTML;
        } else {
            var waitingImg = document.createElement("img");
            waitingImg.src = "/images/waiting.gif";
            waitingImg.alt = "running...";
            div.appendChild(waitingImg);
        }
        searchingEl = div;
    };
    init();
}

/**
 *获取某个HTML Element绝对位置
 *@private
 */
WaitingTip.prototype.GetAbsoluteLocation = function (element) {
    if (arguments.length != 1 || element == null) {
        return null;
    }
    var offsetTop = element.offsetTop;
    var offsetLeft = element.offsetLeft;
    var offsetWidth = element.offsetWidth;
    var offsetHeight = element.offsetHeight;
    while (element = element.offsetParent) {
        offsetTop += element.offsetTop;
        offsetLeft += element.offsetLeft;
    }
    return {
        absoluteTop: offsetTop,
        absoluteLeft: offsetLeft,
        offsetWidth: offsetWidth,
        offsetHeight: offsetHeight
    };
};

/**
 *隐藏
 *@public
 */
WaitingTip.prototype.hide = function () {
    this.getWaitEl().style.display = "none";
};


/**
 *显示
 *@public
 *@param <String> relativelyElId 参照节点的id
 *@param <String>anchor  默认为初始化设置时值
 */
WaitingTip.prototype.show = function (relativelyEl, anchor) {
    var p = this.GetAbsoluteLocation(relativelyEl);
    var waitEl = this.getWaitEl();
    var gap = this.getGap();
    var _anchor = anchor || this.getAnchor();
    waitEl.style.display = "block";
    switch (_anchor) {
        case "down":
            waitEl.style.top = p.absoluteTop + p.offsetHeight + gap + "px";
            waitEl.style.left = p.absoluteLeft + "px";
            break;
        case "right":
            waitEl.style.top = p.absoluteTop + "px";
            waitEl.style.left = p.absoluteLeft + p.offsetWidth + gap + "px";
            break;
        case "left":
            waitElpos = this.GetAbsoluteLocation(waitEl);
            waitEl.style.top = p.absoluteTop + "px";
            waitEl.style.left = p.absoluteLeft - gap - waitElpos.offsetWidth + "px";
            break;
        case "up":
            waitElpos = this.GetAbsoluteLocation(waitEl);
            waitEl.style.top = p.absoluteTop - gap - waitElpos.offsetHeight + "px";
            waitEl.style.left = p.absoluteLeft + "px";
            break;
        case "center":
            try {
                waitElpos = this.GetAbsoluteLocation(waitEl);
                waitEl.style.top = p.absoluteTop + Math.floor((p.offsetHeight - waitElpos.offsetHeight) / 2) + "px";
                waitEl.style.left = p.absoluteLeft + Math.floor((p.offsetWidth - waitElpos.offsetWidth) / 2) + "px";
            } catch (error) {
                waitEl.style.top = p.absoluteTop + Math.floor(p.offsetHeight / 2) + "px";
                waitEl.style.left = p.absoluteLeft + Math.floor(p.offsetWidth / 2) + "px";
            }
            break;
    }
};