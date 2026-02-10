import random
import time
from question_manager import QuestionManager

class PaperGenerator:
    """组卷功能类，负责按知识点随机抽题和试卷生成"""
    
    def __init__(self):
        """
        初始化组卷生成器
        """
        self.question_manager = QuestionManager()
        # 初始化随机种子，确保每次生成的试卷都不同
        random.seed(time.time())
    
    def generate_paper(self, topic=None, question_count=10):
        """
        生成试卷
        
        Args:
            topic: 知识点，None表示所有知识点，列表表示多个知识点
            question_count: 题目数量
            
        Returns:
            tuple: (试卷题目列表, 答案列表)
        """
        # 获取符合条件的题目
        if isinstance(topic, list) and topic:
            # 多个知识点
            questions = []
            for t in topic:
                questions.extend(self.question_manager.get_questions_by_topic(t))
        elif topic:
            # 单个知识点
            questions = self.question_manager.get_questions_by_topic(topic)
        else:
            # 所有知识点
            questions = self.question_manager.get_all_questions()
        
        # 确保题目数量不超过可用题目数
        available_count = len(questions)
        if available_count == 0:
            return [], []
        
        # 调整题目数量
        actual_count = min(question_count, available_count)
        
        # 使用更可靠的随机方法：先创建索引列表，打乱后根据索引选择题目
        # 这样可以确保完全随机，不受原始列表顺序影响
        indices = list(range(len(questions)))
        random.shuffle(indices)
        selected_indices = indices[:actual_count]
        
        # 根据随机索引选择题目
        selected_questions = [questions[i] for i in selected_indices]
        
        # 生成试卷题目列表（不含答案）
        paper_questions = []
        for i, q in enumerate(selected_questions, 1):
            paper_question = {
                'id': i,
                'content': q['content'],
                'topic': q['topic']
            }
            paper_questions.append(paper_question)
        
        # 生成答案列表
        answers = []
        for i, q in enumerate(selected_questions, 1):
            answer = {
                'id': i,
                'content': q['content'],
                'answer': q['answer'],
                'topic': q['topic']
            }
            answers.append(answer)
        
        return paper_questions, answers
    
    def get_available_topics(self):
        """
        获取可用的知识点
        
        Returns:
            list: 知识点列表
        """
        return self.question_manager.get_all_topics()
    
    def get_question_count_by_topic(self, topic=None):
        """
        获取指定知识点的题目数量
        
        Args:
            topic: 知识点，None表示所有知识点
            
        Returns:
            int: 题目数量
        """
        if topic:
            questions = self.question_manager.get_questions_by_topic(topic)
        else:
            questions = self.question_manager.get_all_questions()
        return len(questions)
