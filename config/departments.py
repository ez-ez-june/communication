from __future__ import annotations

from parsers.types import DeptConfig

DEPARTMENTS: list[DeptConfig] = [
    DeptConfig("재정경제부", "https://www.mofe.go.kr/nw/nes/nesdta.do?bbsId=MOSFBBS_000000000030&menuNo=4050100"),
    DeptConfig("과학기술정보통신부", "https://www.msit.go.kr/bbs/list.do?sCode=user&mPid=121&mId=310", include_hint="view.do"),
    DeptConfig("교육부", "https://www.moe.go.kr/boardCnts/listRenew.do?boardID=333&m=020501&s=moe", include_hint="goView", exclude_hint="open.go.kr"),
    DeptConfig("외교부", "https://www.mofa.go.kr/www/brd/m_4075/list.do", include_hint="view.do"),
    DeptConfig("행정안전부", "https://www.mois.go.kr/frt/bbs/type013/commonSelectBoardList.do?bbsId=BBSMSTR_000000000006", include_hint="commonSelectBoardArticle", exclude_hint="type001"),
    DeptConfig("보건복지부", "https://www.mohw.go.kr/board.es?bid=0003&mid=a10501000000", include_hint="act=view"),
    DeptConfig("고용노동부", "https://www.moel.go.kr/news/notice/noticeList.do", include_hint="noticeView"),
    DeptConfig("국토교통부", "https://www.molit.go.kr/USR/BORD0201/m_69/BRD.jsp", include_hint="idx="),
    DeptConfig("통일부", "https://www.unikorea.go.kr/web/unikorea/bbs/bbs_0000000000000001", include_hint="bbs_0000000000000001/"),
    DeptConfig("법무부", "https://www.moj.go.kr/moj/223/subview.do", include_hint="subview.do"),
    DeptConfig("국방부", "https://www.mnd.go.kr/mnd/154/subview.do?enc=Zm5jdDF8QEB8JTJGYmJzJTJGbW5kJTJGMTEwNjYlMkZhcnRjbExpc3QuZG8lM0ZmaW5kT3Bud3JkJTNEJTI2ZmluZFdvcmQlM0QlMjZmaW5kVHlwZSUzRHNqJTI2ZmluZENsU2VxJTNEJTI2X0NTUkYlM0RmZmY2N2Q0NjQ5OTU0MGMwODhkOWJiM2ZiMzM3OGY5ZiUyNg%3D%3D", include_hint="subview.do"),
    DeptConfig("국가보훈부", "https://www.mpva.go.kr/mpva/selectBbsNttList.do?bbsNo=15&key=76", include_hint="selectBbsNttView"),
    DeptConfig("문화체육관광부", "https://www.mcst.go.kr/kor/s_notice/notice/noticeList.jsp", include_hint="noticeView.jsp"),
    DeptConfig("농림축산식품부", "https://www.mafra.go.kr/home/5108/subview.do", include_hint="article"),
    DeptConfig("산업통상부", "https://www.motir.go.kr/kor/article/ATCL6e90bb9de", include_hint="/view?"),
    DeptConfig("기후에너지환경부", "https://mcee.go.kr/home/web/index.do?menuId=10524", include_hint="read.do"),
    DeptConfig("성평등가족부", "https://www.mogef.go.kr/nw/ntc/nw_ntc_s001.do?mid=news400&div1=13&div3=10"),
    DeptConfig("해양수산부", "https://www.mof.go.kr/doc/ko/selectDocList.do?menuSeq=375&bbsSeq=9", include_hint="selectDoc.do"),
    DeptConfig("중소벤처기업부", "https://www.mss.go.kr/site/smba/ex/bbs/List.do?cbIdx=81", include_hint="View.do"),
    DeptConfig("기획예산처", "https://www.mpb.go.kr/web/main/bbs/b_0016/list", include_hint="b_0016/", exclude_hint="list"),
    DeptConfig("인사혁신처", "https://www.mpm.go.kr/mpm/comm/noti/newsNoitice/", include_hint="mode=view"),
    DeptConfig("법제처", "https://www.moleg.go.kr/board.es?mid=a10504000000&bid=0010", include_hint="act=view"),
    DeptConfig("식품의약품안전처", "https://www.mfds.go.kr/brd/m_689/list.do", include_hint="view.do?seq=", exclude_hint="qustnr"),
]

MAJOR_PRESS = [
    {
        "name": "산업통상부",
        "press_url": "https://www.motir.go.kr/kor/article/ATCL3f49a5a8c",
        "include_hint": "/view?",
    },
    {
        "name": "기후에너지환경부",
        "press_url": "https://mcee.go.kr/home/web/index.do?menuId=10523",
        "include_hint": "newsRead",
    },
    {
        "name": "고용노동부",
        "press_url": "https://www.moel.go.kr/news/enews/report/enewsList.do",
        "include_hint": "enewsView",
    },
    {
        "name": "외교부",
        "press_url": "https://www.mofa.go.kr/www/brd/m_4080/list.do",
        "include_hint": "view.do",
    },
]
