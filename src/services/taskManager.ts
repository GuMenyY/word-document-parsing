import { ParseTask, ParseStatus, ParsedDocument } from '../types'

/**
 * 解析任务管理器
 * 使用内存存储任务（生产环境可换用 Redis）
 */
export class TaskManager {
  private tasks: Map<string, ParseTask> = new Map()

  /**
   * 创建新任务
   */
  createTask(fileName?: string, filePath?: string): ParseTask {
    const task: ParseTask = {
      id: this.generateId(),
      status: 'pending',
      fileName,
      filePath,
      createdAt: new Date().toISOString(),
    }
    this.tasks.set(task.id, task)
    return task
  }

  /**
   * 获取任务
   */
  getTask(id: string): ParseTask | undefined {
    return this.tasks.get(id)
  }

  /**
   * 更新任务状态
   */
  updateTaskStatus(id: string, status: ParseStatus, error?: string): boolean {
    const task = this.tasks.get(id)
    if (!task) return false

    task.status = status
    if (error) {
      task.error = error
    }
    if (status === 'completed' || status === 'failed') {
      task.completedAt = new Date().toISOString()
    }

    this.tasks.set(id, task)
    return true
  }

  /**
   * 设置任务结果
   */
  setTaskResult(id: string, result: ParsedDocument): boolean {
    const task = this.tasks.get(id)
    if (!task) return false

    task.result = result
    task.status = 'completed'
    task.completedAt = new Date().toISOString()

    this.tasks.set(id, task)
    return true
  }

  /**
   * 删除任务
   */
  deleteTask(id: string): boolean {
    return this.tasks.delete(id)
  }

  /**
   * 获取所有任务
   */
  getAllTasks(): ParseTask[] {
    return Array.from(this.tasks.values())
  }

  /**
   * 清理过期任务 (超过1小时的)
   */
  cleanupExpiredTasks(): number {
    const oneHourAgo = Date.now() - 60 * 60 * 1000
    let count = 0

    for (const [id, task] of this.tasks) {
      const createdAt = new Date(task.createdAt).getTime()
      if (createdAt < oneHourAgo) {
        this.tasks.delete(id)
        count++
      }
    }

    return count
  }

  /**
   * 生成唯一ID
   */
  private generateId(): string {
    return `task_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`
  }
}

// 导出单例
export const taskManager = new TaskManager()
