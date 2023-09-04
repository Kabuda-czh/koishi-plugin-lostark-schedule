export interface LostarkSchedule {
  id: number
  boss: string
  qq: string
  guildId: string
  name: string
  dps1: number
  dps2: number
  mercy: number
  reason?: string
  days: { [key: string]: boolean }
  joinDate: string
  uploadDate: string
}

export interface ScheduleUser {
  user: LostarkSchedule
  role: 'dps1' | 'dps2' | 'mercy'
}

export interface Round {
  users: ScheduleUser[];
}
