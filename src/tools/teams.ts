/**
 * Microsoft Teams tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { TeamsChannel, TeamsMessage } from '../types.js';

export class TeamsTools {
  constructor(private graphClient: Client, private userId: string) {}

  /**
   * List all teams the user is a member of
   */
  async listTeams(): Promise<any[]> {
    const result = await this.graphClient
      .api('/me/joinedTeams')
      .get();

    return result.value;
  }

  /**
   * Get a specific team
   */
  async getTeam(teamId: string): Promise<any> {
    const team = await this.graphClient
      .api(`/teams/${teamId}`)
      .get();

    return team;
  }

  /**
   * List channels in a team
   */
  async listChannels(teamId: string): Promise<TeamsChannel[]> {
    const result = await this.graphClient
      .api(`/teams/${teamId}/channels`)
      .get();

    return result.value;
  }

  /**
   * Get a specific channel
   */
  async getChannel(teamId: string, channelId: string): Promise<TeamsChannel> {
    const channel = await this.graphClient
      .api(`/teams/${teamId}/channels/${channelId}`)
      .get();

    return channel;
  }

  /**
   * Create a channel
   */
  async createChannel(teamId: string, displayName: string, description?: string): Promise<TeamsChannel> {
    const channel = await this.graphClient
      .api(`/teams/${teamId}/channels`)
      .post({
        displayName,
        description: description || '',
      });

    return channel;
  }

  /**
   * List messages in a channel
   */
  async listChannelMessages(teamId: string, channelId: string, top: number = 50): Promise<TeamsMessage[]> {
    const result = await this.graphClient
      .api(`/teams/${teamId}/channels/${channelId}/messages`)
      .top(top)
      .get();

    return result.value;
  }

  /**
   * Get a specific message
   */
  async getMessage(teamId: string, channelId: string, messageId: string): Promise<TeamsMessage> {
    const message = await this.graphClient
      .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}`)
      .get();

    return message;
  }

  /**
   * Send a message to a channel
   */
  async sendChannelMessage(teamId: string, channelId: string, content: string, contentType: 'html' | 'text' = 'text'): Promise<TeamsMessage> {
    const message = await this.graphClient
      .api(`/teams/${teamId}/channels/${channelId}/messages`)
      .post({
        body: {
          contentType,
          content,
        },
      });

    return message;
  }

  /**
   * Reply to a message
   */
  async replyToMessage(
    teamId: string,
    channelId: string,
    messageId: string,
    content: string,
    contentType: 'html' | 'text' = 'text'
  ): Promise<TeamsMessage> {
    const reply = await this.graphClient
      .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`)
      .post({
        body: {
          contentType,
          content,
        },
      });

    return reply;
  }

  /**
   * List message replies
   */
  async listMessageReplies(teamId: string, channelId: string, messageId: string): Promise<TeamsMessage[]> {
    const result = await this.graphClient
      .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`)
      .get();

    return result.value;
  }

  /**
   * List team members
   */
  async listTeamMembers(teamId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`/teams/${teamId}/members`)
      .get();

    return result.value;
  }

  /**
   * Add member to team
   */
  async addTeamMember(teamId: string, userId: string, roles: string[] = []): Promise<any> {
    const member = await this.graphClient
      .api(`/teams/${teamId}/members`)
      .post({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles,
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`,
      });

    return member;
  }

  /**
   * Remove member from team
   */
  async removeTeamMember(teamId: string, membershipId: string): Promise<void> {
    await this.graphClient
      .api(`/teams/${teamId}/members/${membershipId}`)
      .delete();
  }

  /**
   * List apps installed in a team
   */
  async listTeamApps(teamId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`/teams/${teamId}/installedApps`)
      .expand('teamsAppDefinition')
      .get();

    return result.value;
  }

  /**
   * Send a chat message (1:1 or group chat)
   */
  async sendChatMessage(chatId: string, content: string, contentType: 'html' | 'text' = 'text'): Promise<TeamsMessage> {
    const message = await this.graphClient
      .api(`/chats/${chatId}/messages`)
      .post({
        body: {
          contentType,
          content,
        },
      });

    return message;
  }

  /**
   * List user's chats
   */
  async listChats(): Promise<any[]> {
    const result = await this.graphClient
      .api('/me/chats')
      .get();

    return result.value;
  }

  /**
   * Get online meetings
   */
  async listOnlineMeetings(): Promise<any[]> {
    const result = await this.graphClient
      .api(`/users/${this.userId}/onlineMeetings`)
      .get();

    return result.value;
  }

  /**
   * Create an online meeting
   */
  async createOnlineMeeting(subject: string, startDateTime: string, endDateTime: string): Promise<any> {
    const meeting = await this.graphClient
      .api(`/users/${this.userId}/onlineMeetings`)
      .post({
        subject,
        startDateTime,
        endDateTime,
      });

    return meeting;
  }
}
