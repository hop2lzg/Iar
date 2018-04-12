import urllib
import urllib2
import time
import ssl

class ArcModel:
    def __init__(self):
        # self.logger.debug("import arc")
        self._accpet = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
        self._accept_language = 'zh-CN,zh;q=0.8,en;q=0.6'
        self._cache_control = 'max-age=0'
        self._connection = 'keep-alive'
        self._content_type = 'application/x-www-form-urlencoded'
        self._upgrade_insecure_requests = '1'
        self._user_agent = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36'
        self._cookies = urllib2.HTTPCookieProcessor()
        self._opener = urllib2.build_opener(self._cookies)

    def __try_request(self, req, max_try_num=2):
        res = None
        # res = self._opener.open(req, timeout=0.1)
        for tries in range(max_try_num):
            try:
                # self.logger.debug("Request start at %d times" % tries)
                print "Request start at %d times" % tries
                res = self._opener.open(req, timeout=0.1)
                print "Request success at %d times" % tries
                # self.logger.debug("Request success at %d times" % tries)
                break
            except (urllib2.URLError, ssl.SSLError) as e:
                if tries < (max_try_num - 1):
                    time.sleep(2)
                    continue
                else:
                    print e
            except Exception, e:
                print repr(e)
                break
        return res

    def iar(self):
        # self.logger.debug("GO TO IAR")
        values = {
            'userID': "",
            'user': "muling",
            'password': "lzgsdd"
        }

        url = "https://myarc.arccorp.com/PortalApp/PreLogin.portal"
        data = urllib.urlencode(values)

        headers = {
            'Accept': self._accpet,
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': self._accept_language,
            'Cache-Control': self._cache_control,
            'Connection': self._connection,
            'Content-Length': len(data),
            'Content-Type': self._content_type,
            'Host': 'myarc.arccorp.com',
            'Origin': 'https://myarc.arccorp.com',
            'Referer': 'https://myarc.arccorp.com/PortalApp/PreLogin.portal?logout',
            'Upgrade-Insecure-Requests': self._upgrade_insecure_requests,
            'User-Agent': self._user_agent
        }

        req = urllib2.Request(url, data, headers)
        self.__try_request(req)

arc_model = ArcModel()
arc_model.iar()