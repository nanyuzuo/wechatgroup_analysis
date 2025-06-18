import re
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
import os
import time
import sys
from wxauto import WeChat
import pyperclip
import numpy as np
from matplotlib.path import Path
from matplotlib.patches import PathPatch
import textwrap
from datetime import datetime
import matplotlib.colors as mcolors
from matplotlib.gridspec import GridSpec
from wordcloud import WordCloud

class WeChatGroupAnalyzer:
    def __init__(self):
        self.wx = None
        self.members = []
        self.admin_members = []
        self.province_city_members = {}  # 省份-城市二级结构
        self.foreign_members = []   # 国外成员
        self.unknown_members = []   # 未知地区人员
        self.group_name = ""  # 添加群名属性
        self.initialize_wechat()
        
    def initialize_wechat(self, max_retries=3):
        """初始化微信连接，包含重试机制"""
        print("正在连接微信...")
        for i in range(max_retries):
            try:
                self.wx = WeChat()
                if not self.wx:
                    raise Exception("未找到微信窗口")
                print("微信连接成功！")
                return
            except Exception as e:
                if i < max_retries - 1:
                    print(f"连接失败，正在重试... ({i + 1}/{max_retries})")
                    print("请确保：")
                    print("1. 微信已经打开并登录")
                    print("2. 微信窗口没有被最小化")
                    print("3. 使用微信 3.9.2.23 版本以获得最佳兼容性")
                    time.sleep(2)
                else:
                    print("无法连接到微信，请检查：")
                    print("1. 微信是否正常运行并登录")
                    print("2. 微信窗口是否被最小化")
                    print("3. 微信版本是否兼容（推荐使用 3.9.2.23 版本）")
                    print(f"错误信息: {str(e)}")
                    sys.exit(1)

    def get_group_members(self, group_name):
        """获取微信群成员信息"""
        try:
            print(f"正在搜索群：{group_name}")
            # 切换到群聊
            if not self.wx.ChatWith(group_name):
                raise Exception(f"未找到群：{group_name}")
            
            print("正在获取群成员信息...")
            time.sleep(1)  # 等待群聊窗口完全加载
            
            try:
                # 获取当前聊天窗口的句柄
                chat_window = self.wx.GetSessionList()
                if not chat_window:
                    raise Exception("未能获取到群聊窗口")
                
                # 点击群聊名称
                members = []
                current_members = self.wx.GetGroupMembers()
                if not current_members:
                    print("尝试备选方法获取群成员...")
                    # 尝试使用其他方法获取群成员
                    chat_text = self.wx.GetAllTestData()
                    if chat_text:
                        # 分析聊天记录中的成员信息
                        for line in chat_text:
                            if re.match(r'\d+-[^-]+-.*', line):  # 匹配"学号-城市-昵称"格式
                                members.append(line.strip())
                            elif '马哥' in line:
                                members.append(line.strip())
                else:
                    members = [m.strip() for m in current_members if m.strip()]
                
                if not members:
                    raise Exception("未能获取到群成员信息")
                
                # 去重
                members = list(dict.fromkeys(members))
                
                print(f"成功获取到 {len(members)} 个群成员信息")
                return members
                
            except Exception as e:
                print(f"获取群成员时出错：{str(e)}")
                print("尝试其他方法...")
                
                # 尝试从聊天记录中获取成员信息
                chat_text = self.wx.GetAllTestData()
                members = []
                
                if chat_text:
                    for line in chat_text:
                        # 匹配群成员格式
                        if re.match(r'\d+-[^-]+-.*', line):  # 学号-城市-昵称
                            members.append(line.strip())
                        elif '马哥' in line:  # 马哥教育成员
                            members.append(line.strip())
                
                if not members:
                    raise Exception("未能通过任何方法获取到群成员信息")
                
                # 去重
                members = list(dict.fromkeys(members))
                print(f"通过备选方法获取到 {len(members)} 个群成员信息")
                return members
                
        except Exception as e:
            print(f"获取群成员信息失败：{str(e)}")
            print("请确保：")
            print("1. 群名称输入正确")
            print("2. 您是该群的成员")
            print("3. 使用微信 3.9.11.17 版本以获得最佳兼容性")
            print("4. 群聊窗口处于打开状态")
            print("5. 群聊中有最近的聊天记录")
            print("\n调试信息：")
            print(f"- 当前窗口标题：{self.wx.GetWindowTitle()}")
            print(f"- 会话列表状态：{bool(self.wx.GetSessionList())}")
            sys.exit(1)
    
    def get_location_info(self):
        """获取地理位置信息"""
        return {
            # 直辖市
            '北京': {
                'type': 'municipality',
                'cities': ['北京'],
                'aliases': ['京']
            },
            '上海': {
                'type': 'municipality',
                'cities': ['上海'],
                'aliases': ['沪']
            },
            '天津': {
                'type': 'municipality',
                'cities': ['天津'],
                'aliases': ['津']
            },
            '重庆': {
                'type': 'municipality',
                'cities': ['重庆'],
                'aliases': ['渝']
            },
            
            # 省份及其城市
            '广东': {
                'type': 'province',
                'cities': ['广州', '深圳', '珠海', '汕头', '佛山', '韶关', '湛江', '肇庆', '江门', '茂名', '惠州', '梅州', '汕尾', '河源', '阳江', '清远', '东莞', '中山', '潮州', '揭阳', '云浮'],
                'aliases': ['粤']
            },
            '浙江': {
                'type': 'province',
                'cities': ['杭州', '宁波', '温州', '嘉兴', '湖州', '绍兴', '金华', '衢州', '舟山', '台州', '丽水'],
                'aliases': ['浙']
            },
            '江苏': {
                'type': 'province',
                'cities': ['南京', '无锡', '徐州', '常州', '苏州', '南通', '连云港', '淮安', '盐城', '扬州', '镇江', '泰州', '宿迁'],
                'aliases': ['苏']
            },
            '山东': {
                'type': 'province',
                'cities': ['济南', '青岛', '淄博', '枣庄', '东营', '烟台', '潍坊', '济宁', '泰安', '威海', '日照', '临沂', '德州', '聊城', '滨州', '菏泽'],
                'aliases': ['鲁']
            },
            '河南': {
                'type': 'province',
                'cities': ['郑州', '开封', '洛阳', '平顶山', '安阳', '鹤壁', '新乡', '焦作', '濮阳', '许昌', '漯河', '三门峡', '南阳', '商丘', '信阳', '周口', '驻马店'],
                'aliases': ['豫']
            },
            '湖北': {
                'type': 'province',
                'cities': ['武汉', '黄石', '十堰', '宜昌', '襄阳', '鄂州', '荆门', '孝感', '荆州', '黄冈', '咸宁', '随州', '恩施'],
                'aliases': ['鄂']
            },
            '湖南': {
                'type': 'province',
                'cities': ['长沙', '株洲', '湘潭', '衡阳', '邵阳', '岳阳', '常德', '张家界', '益阳', '郴州', '永州', '怀化', '娄底', '湘西'],
                'aliases': ['湘']
            },
            '河北': {
                'type': 'province',
                'cities': ['石家庄', '唐山', '秦皇岛', '邯郸', '邢台', '保定', '张家口', '承德', '沧州', '廊坊', '衡水'],
                'aliases': ['冀']
            },
            '山西': {
                'type': 'province',
                'cities': ['太原', '大同', '阳泉', '长治', '晋城', '朔州', '晋中', '运城', '忻州', '临汾', '吕梁'],
                'aliases': ['晋']
            },
            '内蒙古': {
                'type': 'autonomous',
                'cities': ['呼和浩特', '包头', '乌海', '赤峰', '通辽', '鄂尔多斯', '呼伦贝尔', '巴彦淖尔', '乌兰察布'],
                'aliases': ['内蒙古自治区', '内蒙']
            },
            '辽宁': {
                'type': 'province',
                'cities': ['沈阳', '大连', '鞍山', '抚顺', '本溪', '丹东', '锦州', '营口', '阜新', '辽阳', '盘锦', '铁岭', '朝阳', '葫芦岛'],
                'aliases': ['辽']
            },
            '吉林': {
                'type': 'province',
                'cities': ['长春', '吉林', '四平', '辽源', '通化', '白山', '松原', '白城', '延边'],
                'aliases': ['吉']
            },
            '黑龙江': {
                'type': 'province',
                'cities': ['哈尔滨', '齐齐哈尔', '鸡西', '鹤岗', '双鸭山', '大庆', '伊春', '佳木斯', '七台河', '牡丹江', '黑河', '绥化', '大兴安岭'],
                'aliases': ['黑']
            },
            '陕西': {
                'type': 'province',
                'cities': ['西安', '铜川', '宝鸡', '咸阳', '渭南', '延安', '汉中', '榆林', '安康', '商洛'],
                'aliases': ['陕']
            },
            '甘肃': {
                'type': 'province',
                'cities': ['兰州', '嘉峪关', '金昌', '白银', '天水', '武威', '张掖', '平凉', '酒泉', '庆阳', '定西', '陇南', '临夏', '甘南'],
                'aliases': ['甘']
            },
            '青海': {
                'type': 'province',
                'cities': ['西宁', '海东', '海北', '黄南', '海南', '果洛', '玉树', '海西'],
                'aliases': ['青']
            },
            '宁夏': {
                'type': 'autonomous',
                'cities': ['银川', '石嘴山', '吴忠', '固原', '中卫'],
                'aliases': ['宁夏回族自治区', '宁']
            },
            '新疆': {
                'type': 'autonomous',
                'cities': ['乌鲁木齐', '克拉玛依', '吐鲁番', '哈密', '昌吉', '博尔塔拉', '巴音郭楞', '阿克苏', '克孜勒苏', '喀什', '和田', '伊犁', '塔城', '阿勒泰'],
                'aliases': ['新疆维吾尔自治区', '新']
            },
            '四川': {
                'type': 'province',
                'cities': ['成都', '自贡', '攀枝花', '泸州', '德阳', '绵阳', '广元', '遂宁', '内江', '乐山', '南充', '眉山', '宜宾', '广安', '达州', '雅安', '巴中', '资阳', '阿坝', '甘孜', '凉山'],
                'aliases': ['川']
            },
            '贵州': {
                'type': 'province',
                'cities': ['贵阳', '六盘水', '遵义', '安顺', '毕节', '铜仁', '黔西南', '黔东南', '黔南'],
                'aliases': ['贵']
            },
            '云南': {
                'type': 'province',
                'cities': ['昆明', '曲靖', '玉溪', '保山', '昭通', '丽江', '普洱', '临沧', '楚雄', '红河', '文山', '西双版纳', '大理', '德宏', '怒江', '迪庆'],
                'aliases': ['云']
            },
            '西藏': {
                'type': 'autonomous',
                'cities': ['拉萨', '日喀则', '昌都', '林芝', '山南', '那曲', '阿里'],
                'aliases': ['西藏自治区', '藏']
            },
            '安徽': {
                'type': 'province',
                'cities': ['合肥', '芜湖', '蚌埠', '淮南', '马鞍山', '淮北', '铜陵', '安庆', '黄山', '滁州', '阜阳', '宿州', '六安', '亳州', '池州', '宣城'],
                'aliases': ['皖']
            },
            '江西': {
                'type': 'province',
                'cities': ['南昌', '景德镇', '萍乡', '九江', '新余', '鹰潭', '赣州', '吉安', '宜春', '抚州', '上饶'],
                'aliases': ['赣']
            },
            '福建': {
                'type': 'province',
                'cities': ['福州', '厦门', '莆田', '三明', '泉州', '漳州', '南平', '龙岩', '宁德'],
                'aliases': ['闽']
            },
            '广西': {
                'type': 'autonomous',
                'cities': ['南宁', '柳州', '桂林', '梧州', '北海', '防城港', '钦州', '贵港', '玉林', '百色', '贺州', '河池', '来宾', '崇左'],
                'aliases': ['广西壮族自治区', '桂']
            },
            '海南': {
                'type': 'province',
                'cities': ['海口', '三亚', '三沙', '儋州'],
                'aliases': ['琼']
            },
            
            # 特别行政区
            '香港': {
                'type': 'special',
                'cities': ['香港'],
                'aliases': ['港']
            },
            '澳门': {
                'type': 'special',
                'cities': ['澳门'],
                'aliases': ['澳']
            },
            '台湾': {
                'type': 'special',
                'cities': ['台北', '高雄', '台中', '台南', '新北'],
                'aliases': ['台']
            }
        }
        
    def analyze_members(self, members):
        """分析成员信息"""
        location_info = self.get_location_info()
        # 创建城市到省份的映射
        city_to_province = {}
        for province, info in location_info.items():
            for city in info['cities']:
                city_to_province[city] = province
        
        # 国外城市列表
        foreign_cities = {
            '多伦多', '温哥华', '蒙特利尔', '渥太华',  # 加拿大
            '纽约', '洛杉矶', '芝加哥', '休斯顿', '西雅图',  # 美国
            '伦敦', '曼彻斯特', '利物浦',  # 英国
            '巴黎', '马赛', '里昂',  # 法国
            '柏林', '慕尼黑', '汉堡',  # 德国
            '东京', '大阪', '名古屋',  # 日本
            '首尔', '釜山', '仁川',  # 韩国
            '新加坡',  # 新加坡
            '悉尼', '墨尔本', '布里斯班',  # 澳大利亚
            '迪拜', '阿布扎比',  # 阿联酋
            '莫斯科', '圣彼得堡',  # 俄罗斯
            '马德里', '巴塞罗那',  # 西班牙
            '米兰', '罗马', '威尼斯'  # 意大利
        }
        
        # 初始化分类存储
        self.admin_members = []  # 马哥教育成员
        self.province_city_members = {}  # 按省份和城市分类的成员
        self.foreign_members = []  # 国外成员
        self.unknown_members = []  # 未知分类成员
        
        for member in members:
            # 更严格的空格和不可见字符处理
            member = ''.join(c for c in member if c.isprintable())  # 移除所有不可打印字符
            member = ' '.join(part.strip() for part in member.split())  # 分割并重组，确保只有单个空格
            
            # 判断马哥教育成员
            if '马哥' in member or '班' in member or '豆' in member:
                self.admin_members.append(member)
                continue
            
            # 尝试匹配地理位置信息
            location_found = False
            member_parts = member.split('-')
            
            # 1. 优先尝试匹配城市（因为城市信息更具体）
            matched_city = None
            for city in sorted(city_to_province.keys(), key=len, reverse=True):  # 按城市名长度降序排序，避免"新"匹配到"新疆"
                if city in member:
                    matched_city = city
                    province = city_to_province[city]
                    
                    # 初始化省份数据结构（如果不存在）
                    if province not in self.province_city_members:
                        self.province_city_members[province] = {'total': 0, 'cities': {}}
                    
                    # 初始化城市数据结构（如果不存在）
                    if city not in self.province_city_members[province]['cities']:
                        self.province_city_members[province]['cities'][city] = []
                    
                    # 添加成员到对应的城市
                    self.province_city_members[province]['cities'][city].append(member)
                    self.province_city_members[province]['total'] += 1
                    location_found = True
                    break
            
            # 2. 如果没有匹配到城市，尝试匹配省份
            if not location_found and len(member_parts) > 1:
                location_part = member_parts[1].strip()
                for province, info in location_info.items():
                    if location_part == province or location_part in info['aliases']:
                        if province not in self.province_city_members:
                            self.province_city_members[province] = {'total': 0, 'cities': {}}
                        if '省会' not in self.province_city_members[province]['cities']:
                            self.province_city_members[province]['cities']['省会'] = []
                        self.province_city_members[province]['cities']['省会'].append(member)
                        self.province_city_members[province]['total'] += 1
                        location_found = True
                        break
            
            # 3. 检查是否是国外城市
            if not location_found:
                for city in foreign_cities:
                    if city in member:
                        self.foreign_members.append(member)
                        location_found = True
                        break
            
            # 4. 如果仍然没有匹配到，归类到未知
            if not location_found:
                self.unknown_members.append(member)
        
        # 输出分析结果
        print(f"\n分析结果：")
        print(f"总成员数：{len(members)}")
        print(f"马哥教育成员数：{len(self.admin_members)}")
        
        # 打印马哥教育成员详情
        if self.admin_members:
            print("\n马哥教育成员详情：")
            for member in self.admin_members:
                print(f"- {member}")
        
        # 打印省份和城市分布详情
        if self.province_city_members:
            print("\n各省份及城市分布详情：")
            # 按总人数降序排序省份
            sorted_provinces = sorted(self.province_city_members.items(), 
                                   key=lambda x: (-x[1]['total'], x[0]))
            
            for province, data in sorted_provinces:
                print(f"\n{province}（共{data['total']}人）：")
                # 按人数降序排序城市
                sorted_cities = sorted(data['cities'].items(), 
                                    key=lambda x: (-len(x[1]), x[0]))
                
                # 首先显示未知城市的成员（如果有）
                if '省会' in data['cities'] and data['cities']['省会']:
                    print(f"- {province}未知城市（{len(data['cities']['省会'])}人）")
                    for member in data['cities']['省会']:
                        print(f"  * {member}")
                
                # 然后显示各个城市的成员
                for city, city_members in sorted_cities:
                    if city != '省会':
                        print(f"- {city}（{len(city_members)}人）")
                        for member in city_members:
                            print(f"  * {member}")
        
        # 打印国外成员详情
        if self.foreign_members:
            print("\n国外成员详情：")
            print(f"共计：{len(self.foreign_members)}人")
            for member in self.foreign_members:
                print(f"- {member}")
        
        # 打印未知分类成员
        if self.unknown_members:
            print("\n未知地区人员：")
            print(f"共计：{len(self.unknown_members)}人")
            for member in self.unknown_members:
                print(f"- {member}")
    
    def generate_text_result(self):
        """生成文本统计结果"""
        result = []
        
        # 添加总体统计
        result.append("=== 微信群成员分析报告 ===\n")
        
        # 添加综述段落
        total_members = (len(self.admin_members) + 
                       sum(data['total'] for data in self.province_city_members.values()) +
                       len(self.foreign_members) + 
                       len(self.unknown_members))
        result.append(f"该群共有成员{total_members}人，具体构成如下：\n")
        
        # 添加马哥教育管理员信息
        result.append(f"【马哥教育成员】（{len(self.admin_members)}人）")
        for member in self.admin_members:
            result.append(f"- {member}")
        result.append("")
        
        # 添加省份和城市分布信息
        result.append("【地区分布情况】")
        
        # 按总人数降序排序省份
        sorted_provinces = sorted(self.province_city_members.items(), 
                               key=lambda x: (-x[1]['total'], x[0]))
        
        for province, data in sorted_provinces:
            result.append(f"\n{province}（共{data['total']}人）：")
            # 按人数降序排序城市
            sorted_cities = sorted(data['cities'].items(), 
                                key=lambda x: (-len(x[1]), x[0]))
            
            # 首先显示未知城市的成员（如果有）
            if '省会' in data['cities'] and data['cities']['省会']:
                result.append(f"- {province}未知城市（{len(data['cities']['省会'])}人）")
                for member in data['cities']['省会']:
                    result.append(f"  * {member}")
            
            # 然后显示各个城市的成员
            for city, city_members in sorted_cities:
                if city != '省会':
                    result.append(f"- {city}（{len(city_members)}人）")
                    for member in city_members:
                        result.append(f"  * {member}")
        
        # 添加国外成员信息
        if self.foreign_members:
            result.append(f"\n【国外成员】（{len(self.foreign_members)}人）")
            for member in self.foreign_members:
                result.append(f"- {member}")
        
        # 添加未知分类成员
        if self.unknown_members:
            result.append(f"\n【未知地区人员】（{len(self.unknown_members)}人）")
            for member in self.unknown_members:
                result.append(f"- {member}")
            
        return "\n".join(result)
    
    def create_custom_marker(self):
        """创建自定义人形图标"""
        # 创建人形图标的顶点
        person = np.array([
            # 头部（圆形）
            [0.0, 0.7],  [0.1, 0.8],  [0.2, 0.85], [0.3, 0.8],
            [0.4, 0.7],  [0.3, 0.6],  [0.2, 0.55], [0.1, 0.6],
            [0.0, 0.7],
            # 身体
            [0.2, 0.55], [0.2, 0.2],
            # 手臂
            [0.0, 0.4],  [0.2, 0.4],  [0.4, 0.4],
            # 回到身体
            [0.2, 0.2],
            # 腿部
            [0.1, 0.0],  [0.2, 0.2],  [0.3, 0.0]
        ])
        
        # 将图标缩放到合适大小
        person = person - [0.2, 0.4]  # 将图标中心移到原点
        person = person * 0.5         # 缩放大小
        
        return person

    def create_charts(self):
        """创建统计图表"""
        plt.rcParams['font.sans-serif'] = ['SimHei']  # 设置中文字体
        plt.rcParams['axes.unicode_minus'] = False    # 解决负号显示问题
        
        # 创建图表
        fig, ax = plt.subplots(figsize=(15, 10))
        
        # 准备数据
        categories = []
        counts = []
        
        # 1. 马哥教育成员
        categories.append('马哥教育成员')
        counts.append(len(self.admin_members))
        
        # 2. 各省份成员（按人数降序）
        province_data = []
        for province, data in self.province_city_members.items():
            province_data.append((province, data['total']))
        
        # 按人数降序排列省份
        province_data.sort(key=lambda x: x[1], reverse=True)
        
        for province, count in province_data:
            categories.append(province)
            counts.append(count)
        
        # 3. 国外成员
        categories.append('国外成员')
        counts.append(len(self.foreign_members))
        
        # 4. 未知地区人员
        categories.append('未知地区人员')
        counts.append(len(self.unknown_members))
        
        # 创建水平条形图
        y_pos = np.arange(len(categories))
        
        # 设置统一的浅蓝色，马哥教育成员使用红色
        colors = ['#FF6B6B' if i == 0 else '#5B9BD5' for i in range(len(categories))]
        bars = ax.barh(y_pos, counts, color=colors)
        
        # 设置图表样式
        ax.set_yticks(y_pos)
        ax.set_yticklabels(categories, fontsize=10)
        ax.set_xlabel('人数', fontsize=12)
        
        # 设置背景色和网格样式
        ax.set_facecolor('#FFFFFF')
        fig.patch.set_facecolor('#FFFFFF')
        ax.grid(axis='x', linestyle='--', alpha=0.3, color='#E5E5E5')
        
        # 在条形图右端添加数值标签
        for i, bar in enumerate(bars):
            width = bar.get_width()
            ax.text(width + 0.5, bar.get_y() + bar.get_height()/2,
                   f'{int(width)}人',
                   ha='left', va='center', fontsize=10)
        
        # 计算总人数
        total_members = sum(counts)
        
        # 设置标题和总人数说明
        plt.title('成员分布情况\n本群共{}人，构成情况如图'.format(total_members), 
                 pad=20, fontsize=14, loc='center')
        
        # 调整布局
        plt.tight_layout()
        
        # 保存图表
        plt.savefig('statistics_charts.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print("统计图表已生成：statistics_charts.png")

    def create_text_image(self, text, width=1200, font_size=24):
        """将文本转换为图片"""
        # 设置字体
        font = ImageFont.truetype("simhei.ttf", font_size)
        font_small = ImageFont.truetype("simhei.ttf", font_size - 4)
        
        # 计算行高和边距
        line_height = font_size * 1.5
        padding = 40
        
        # 分割文本行
        lines = text.split('\n')
        
        # 计算每行实际宽度和换行
        wrapped_lines = []
        for line in lines:
            # 缩进处理
            indent = len(line) - len(line.lstrip())
            indent_space = "  " * indent
            
            # 处理实际内容
            content = line.lstrip()
            
            # 根据内容类型设置字体和颜色
            if content.startswith('==='):  # 主标题
                current_font = ImageFont.truetype("simhei.ttf", font_size + 4)
                color = '#FF6B6B'  # 主题色
            elif content.startswith('【'):  # 分类标题
                current_font = ImageFont.truetype("simhei.ttf", font_size + 2)
                color = '#4ECDC4'  # 次要主题色
            elif content.startswith('- '):  # 一级列表项
                current_font = font
                color = '#2C3E50'  # 深灰色
            elif content.startswith('* '):  # 二级列表项
                current_font = font_small
                color = '#5D6D7E'  # 中灰色
            else:  # 普通文本
                current_font = font
                color = '#34495E'  # 标准文本色
            
            # 处理缩进和行宽
            if content.startswith('*'):
                line_width = width - padding * 3
            else:
                line_width = width - padding * 2
            
            # 添加到行列表
            wrapped_lines.append((indent_space + content, current_font, color))
        
        # 计算所需图片高度
        height = int(len(wrapped_lines) * line_height + padding * 2)
        
        # 创建图片
        image = Image.new('RGB', (width, height), '#FFFFFF')  # 纯白背景
        draw = ImageDraw.Draw(image)
        
        # 绘制文本
        y = padding
        for line, line_font, color in wrapped_lines:
            draw.text((padding, y), line, font=line_font, fill=color)
            y += line_height
        
        return image

    def merge_images(self, text_image, chart_image):
        """合并文本图片和统计图表"""
        # 创建标题图片
        title_height = 100
        title_image = Image.new('RGB', (text_image.width, title_height), '#FFFFFF')
        draw = ImageDraw.Draw(title_image)
        
        # 设置标题字体
        title_font = ImageFont.truetype("simhei.ttf", 36)
        
        # 绘制标题（居中）
        title_text = "马哥大模型1期成员构成分析报告"
        title_bbox = draw.textbbox((0, 0), title_text, font=title_font)
        title_width = title_bbox[2] - title_bbox[0]
        x = (text_image.width - title_width) // 2
        draw.text((x, 30), title_text, font=title_font, fill='#2C3E50')
        
        # 调整统计图表大小以匹配文本宽度
        chart_width = text_image.width
        chart_height = int(chart_image.height * (chart_width / chart_image.width))
        chart_image = chart_image.resize((chart_width, chart_height), Image.Resampling.LANCZOS)
        
        # 创建新图片（标题 + 图表 + 文本）
        new_height = title_height + chart_height + text_image.height
        merged_image = Image.new('RGB', (chart_width, new_height), '#FFFFFF')
        
        # 按顺序粘贴图片：标题 -> 图表 -> 文本
        merged_image.paste(title_image, (0, 0))
        merged_image.paste(chart_image, (0, title_height))
        merged_image.paste(text_image, (0, title_height + chart_height))
        
        return merged_image

    def generate_report(self):
        """生成完整的分析报告"""
        # 生成文本报告
        text_content = self.generate_text_result()
        
        # 保存文本报告
        with open('group_analysis.txt', 'w', encoding='utf-8') as f:
            f.write(text_content)
        
        # 生成统计图表
        self.create_charts()
        
        # 将文本转换为图片
        text_image = self.create_text_image(text_content)
        
        # 读取统计图表
        chart_image = Image.open('statistics_charts.png')
        
        # 合并图片
        final_image = self.merge_images(text_image, chart_image)
        
        # 保存最终图片
        final_image.save('group_analysis.png', quality=95, dpi=(300, 300))
        
        # 删除临时的统计图表文件
        os.remove('statistics_charts.png')
        
        print("分析完成！生成的文件：")
        print("1. group_analysis.png - 完整的图片格式分析报告")
        print("2. group_analysis.txt - 文本格式统计结果")

    def run(self):
        """运行分析器"""
        # 获取要分析的群名称
        group_name = input("请输入要分析的微信群名称：")
        
        # 分析群成员
        self.analyze_members(self.get_group_members(group_name))
        
        # 生成报告
        self.generate_report()

class ModernUIGenerator:
    def __init__(self, width=1200, height=2000):
        self.width = width
        self.height = height
        self.background_color = '#F5F7FA'
        self.primary_color = '#FF6B6B'
        self.secondary_color = '#4ECDC4'
        self.text_color = '#2C3E50'
        self.font_path = "simhei.ttf"
        
    def create_gradient_background(self):
        """创建渐变背景"""
        image = Image.new('RGB', (self.width, self.height), self.background_color)
        draw = ImageDraw.Draw(image)
        
        # 创建顶部渐变条
        for y in range(60):
            color = self._interpolate_color(
                self.primary_color, 
                self.secondary_color, 
                y/60
            )
            draw.line([(0, y), (self.width, y)], fill=color)
            
        return image
        
    def _interpolate_color(self, color1, color2, factor):
        """在两个颜色之间进行插值"""
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            
        def rgb_to_hex(rgb):
            return '#{:02x}{:02x}{:02x}'.format(*rgb)
            
        c1 = hex_to_rgb(color1)
        c2 = hex_to_rgb(color2)
        
        rgb = tuple(int(c1[i] + (c2[i] - c1[i]) * factor) for i in range(3))
        return rgb_to_hex(rgb)
        
    def create_card(self, x, y, width, height, title, value, icon=None):
        """创建数据卡片"""
        card = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(card)
        
        # 添加卡片阴影
        shadow = Image.new('RGBA', (width, height), (0, 0, 0, 0))
        shadow_draw = ImageDraw.Draw(shadow)
        shadow_draw.rectangle([0, 0, width, height], fill=(0, 0, 0, 30))
        
        # 绘制卡片内容
        title_font = ImageFont.truetype(self.font_path, 20)
        value_font = ImageFont.truetype(self.font_path, 36)
        
        # 绘制标题
        draw.text((20, 15), title, font=title_font, fill=self.text_color)
        
        # 绘制数值
        draw.text((20, 45), str(value), font=value_font, fill=self.primary_color)
        
        return card, shadow
        
    def create_time_chart(self, data, width, height):
        """创建24小时活跃度图表"""
        fig, ax = plt.subplots(figsize=(width/100, height/100), dpi=100)
        
        # 使用渐变色填充
        colors = [self.primary_color, self.secondary_color]
        cmap = mcolors.LinearSegmentedColormap.from_list("", colors)
        
        # 绘制图表
        ax.plot(data.index, data.values, color=self.primary_color)
        ax.fill_between(data.index, data.values, alpha=0.3, color=self.secondary_color)
        
        # 设置样式
        ax.set_facecolor(self.background_color)
        fig.patch.set_facecolor(self.background_color)
        
        plt.title("24小时活跃度分析", pad=20, fontsize=14, color=self.text_color)
        plt.grid(True, alpha=0.3)
        
        return fig

    def create_word_cloud(self, text_data, width, height):
        """创建词云图"""
        wordcloud = WordCloud(
            width=width,
            height=height,
            background_color=self.background_color,
            font_path=self.font_path,
            colormap='RdYlBu',
            max_words=100
        ).generate(text_data)
        
        return wordcloud.to_image()

    def create_timeline(self, events, width, height):
        """创建时间轴"""
        timeline = Image.new('RGB', (width, height), self.background_color)
        draw = ImageDraw.Draw(timeline)
        
        # 设置字体
        font = ImageFont.truetype(self.font_path, 16)
        time_font = ImageFont.truetype(self.font_path, 14)
        
        y = 30
        for event in events:
            # 绘制时间点
            draw.ellipse([20, y-5, 30, y+5], fill=self.primary_color)
            
            # 绘制连接线
            if y < height - 50:
                draw.line([25, y+5, 25, y+45], fill=self.primary_color)
            
            # 绘制事件文本
            draw.text((40, y-10), event['time'], font=time_font, fill=self.text_color)
            draw.text((40, y+10), event['content'], font=font, fill=self.text_color)
            
            y += 50
            
        return timeline

def main():
    # 初始化分析器
    analyzer = WeChatGroupAnalyzer()
    
    # 运行分析器
    analyzer.run()

if __name__ == "__main__":
    main() 