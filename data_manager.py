import json
import os

class DataManager:
    """数据管理类，负责JSON文件的读写操作"""
    
    def __init__(self, file_path='questions.json'):
        """
        初始化数据管理器
        
        Args:
            file_path: JSON文件路径
        """
        self.file_path = file_path
        self.data = self._load_data()
    
    def _load_data(self):
        """
        加载JSON文件数据
        
        Returns:
            dict: 题目数据字典
        """
        if not os.path.exists(self.file_path):
            # 如果文件不存在，创建默认数据结构，包含默认题库
            default_data = {
                'banks': [
                    {
                        'id': 1,
                        'name': '题库一',
                        'questions': []
                    }
                ]
            }
            self._save_data(default_data)
            return default_data
        
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 兼容旧数据结构
                if 'questions' in data and 'banks' not in data:
                    # 将旧数据转换为新结构
                    new_data = {
                        'banks': [
                            {
                                'id': 1,
                                'name': '题库一',
                                'questions': data['questions']
                            }
                        ]
                    }
                    self._save_data(new_data)
                    return new_data
                return data
        except json.JSONDecodeError:
            # 如果文件损坏，返回默认数据结构
            return {
                'banks': [
                    {
                        'id': 1,
                        'name': '题库一',
                        'questions': []
                    }
                ]
            }
    
    def _save_data(self, data):
        """
        保存数据到JSON文件
        
        Args:
            data: 要保存的数据
        """
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存数据失败: {e}")
    
    def get_banks(self):
        """
        获取所有题库
        
        Returns:
            list: 题库列表
        """
        return self.data.get('banks', [])
    
    def get_questions(self, bank_id=None):
        """
        获取题目
        
        Args:
            bank_id: 题库ID，None表示所有题库
            
        Returns:
            list: 题目列表
        """
        if bank_id:
            # 获取指定题库的题目
            for bank in self.data['banks']:
                if bank['id'] == bank_id:
                    return bank.get('questions', [])
            return []
        else:
            # 获取所有题库的题目
            all_questions = []
            for bank in self.data['banks']:
                all_questions.extend(bank.get('questions', []))
            return all_questions
    
    def save_questions(self, questions, bank_id):
        """
        保存题目列表
        
        Args:
            questions: 题目列表
            bank_id: 题库ID
        """
        for bank in self.data['banks']:
            if bank['id'] == bank_id:
                bank['questions'] = questions
                self._save_data(self.data)
                break
    
    def get_question_by_id(self, question_id):
        """
        根据ID获取题目
        
        Args:
            question_id: 题目ID
            
        Returns:
            tuple: (题目字典, 题库ID)，如果不存在返回(None, None)
        """
        for bank in self.data['banks']:
            for question in bank.get('questions', []):
                if question.get('id') == question_id:
                    return question, bank['id']
        return None, None
    
    def add_question(self, question, bank_id):
        """
        添加新题目
        
        Args:
            question: 题目字典
            bank_id: 题库ID
        """
        # 生成新ID（全局唯一）
        max_id = 0
        for bank in self.data['banks']:
            for q in bank.get('questions', []):
                if q.get('id', 0) > max_id:
                    max_id = q.get('id', 0)
        question['id'] = max_id + 1
        
        # 添加到指定题库
        for bank in self.data['banks']:
            if bank['id'] == bank_id:
                bank.get('questions', []).append(question)
                self._save_data(self.data)
                break
    
    def update_question(self, question_id, updated_question):
        """
        更新题目
        
        Args:
            question_id: 题目ID
            updated_question: 更新后的题目字典
        """
        for bank in self.data['banks']:
            for i, question in enumerate(bank.get('questions', [])):
                if question.get('id') == question_id:
                    updated_question['id'] = question_id
                    bank.get('questions', [])[i] = updated_question
                    self._save_data(self.data)
                    return True
        return False
    
    def delete_question(self, question_id):
        """
        删除题目
        
        Args:
            question_id: 题目ID
        """
        for bank in self.data['banks']:
            bank['questions'] = [
                q for q in bank.get('questions', []) 
                if q.get('id') != question_id
            ]
        self._save_data(self.data)
    
    def get_unique_topics(self, bank_id=None):
        """
        获取所有唯一的知识点
        
        Args:
            bank_id: 题库ID，None表示所有题库
            
        Returns:
            list: 知识点列表
        """
        topics = set()
        if bank_id:
            # 获取指定题库的知识点
            for bank in self.data['banks']:
                if bank['id'] == bank_id:
                    for question in bank.get('questions', []):
                        if 'topic' in question:
                            topics.add(question['topic'])
                    break
        else:
            # 获取所有题库的知识点
            for bank in self.data['banks']:
                for question in bank.get('questions', []):
                    if 'topic' in question:
                        topics.add(question['topic'])
        return sorted(list(topics))
    
    def add_bank(self, bank_name):
        """
        添加新题库
        
        Args:
            bank_name: 题库名称
            
        Returns:
            int: 新题库的ID
        """
        # 生成新ID
        max_id = 0
        for bank in self.data['banks']:
            if bank['id'] > max_id:
                max_id = bank['id']
        new_id = max_id + 1
        
        # 添加新题库
        new_bank = {
            'id': new_id,
            'name': bank_name,
            'questions': []
        }
        self.data['banks'].append(new_bank)
        self._save_data(self.data)
        return new_id
    
    def update_bank_name(self, bank_id, new_name):
        """
        更新题库名称
        
        Args:
            bank_id: 题库ID
            new_name: 新题库名称
        """
        for bank in self.data['banks']:
            if bank['id'] == bank_id:
                bank['name'] = new_name
                self._save_data(self.data)
                break
    
    def delete_bank(self, bank_id):
        """
        删除题库
        
        Args:
            bank_id: 题库ID
        """
        # 确保至少保留一个题库
        if len(self.data['banks']) > 1:
            self.data['banks'] = [
                bank for bank in self.data['banks'] 
                if bank['id'] != bank_id
            ]
            self._save_data(self.data)
