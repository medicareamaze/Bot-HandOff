import * as builder from 'botbuilder';
import * as bluebird from 'bluebird';
import * as request from 'request';
import * as _ from 'lodash';
import mongoose = require('mongoose');
mongoose.Promise = bluebird;

import { By, Conversation, Provider, ConversationState,Lead } from './handoff';

const indexExports = require('./index');

// -------------------
// Bot Framework types
// -------------------
export const IIdentitySchema = new mongoose.Schema({
    id: { type: String, required: true },
    isGroup: { type: Boolean, required: false },
    name: { type: String, required: false },
}, {
        _id: false,
        strict: false,
    });

export const IAddressSchema = new mongoose.Schema({
    bot: { type: IIdentitySchema, required: true },
    channelId: { type: String, required: true },
    conversation: { type: IIdentitySchema, required: false },
    user: { type: IIdentitySchema, required: true },
    id: { type: String, required: false },
    serviceUrl: { type: String, required: false },
    useAuth: { type: Boolean, required: false }
}, {
        strict: false,
        id: false,
        _id: false
    });

// -------------
// Handoff types
// -------------
export const TranscriptLineSchema = new mongoose.Schema({
    timestamp: {},
    from: String,
    sentimentScore: Number,
    state: Number,
    attachments: String,
    adaptiveResponseKVPairs: String,
    text: String
});

export const ConversationSchema = new mongoose.Schema({
    customer: { type: IAddressSchema, required: true },
    agent: { type: IAddressSchema, required: false },
    state: {
        type: Number,
        required: true,
        min: 0,
        max: 3
    },
    transcript: [TranscriptLineSchema]
});

export interface ConversationDocument extends Conversation, mongoose.Document { }
export const ConversationModel = mongoose.model<ConversationDocument>('Conversation', ConversationSchema)

export const BySchema = new mongoose.Schema({
    bestChoice: Boolean,
    agentConversationId: String,
    customerConversationId: String,
    customerName: String,
    customerId: String
});
export interface ByDocument extends By, mongoose.Document { }
export const ByModel = mongoose.model<ByDocument>('By', BySchema);


export const LeadSchema = new mongoose.Schema({
    leadId: String,
    name: String,
    email:String,
    mobileNumber:String,
    landLine:String,
    zip:String,
    dateOfBirth:Date,
    eligibleProductTypes: [String],
    interestedProductTypes: [String],
    offeredProducts:[String],
    interestedProducts:[String],
    webPushSubscription:[String],
    androidPushSubscription:[String],
    iosPushSubscription:[String],
    isAgent:Boolean,    
    lastConversationsByChannel: [ConversationSchema]
});

export interface LeadDocument extends Lead, mongoose.Document { }
export const LeadModel = mongoose.model<LeadDocument>('Lead', LeadSchema);

export { mongoose };



// -----------------
// Mongoose Provider
// -----------------
export class MongooseProvider implements Provider {
    public init(): void { }
    
    async addToTranscript(by: By, message: builder.IMessage, from: string): Promise<boolean> {
        let sentimentScore = -1;
        let text = message.text;        
        let adaptiveResponseKVPairs = !!message.value? JSON.stringify(message.value):null;
        let attachments = JSON.stringify(message.attachments);
        let datetime = new Date().toISOString();
        let conversation: Conversation = await this.getConversation(by);

        if (!conversation) return false;
       // if(message.user) conversation.customer.user = message.user;
         
        if (from == "Customer") {
            if (indexExports._textAnalyticsKey) { sentimentScore = await this.collectSentiment(text); }
            datetime = message.localTimestamp ? message.localTimestamp : message.timestamp
        }

        conversation.transcript.push({
            timestamp: datetime,
            from: from,
            sentimentScore: sentimentScore,
            state: conversation.state,
            attachments: attachments,
            adaptiveResponseKVPairs:adaptiveResponseKVPairs,
            text
        });

        if (indexExports._appInsights) {   
            // You can't log embedded json objects in application insights, so we are flattening the object to one item.
            // Also, have to stringify the object so functions from mongodb don't get logged 
            let latestTranscriptItem = conversation.transcript.length-1;
            let x = JSON.parse(JSON.stringify(conversation.transcript[latestTranscriptItem]));
            x['botId'] = conversation.customer.bot.id;
            x['customerId'] = conversation.customer.user.id;
            x['customerName'] = conversation.customer.user.name;
            x['customerChannelId'] = conversation.customer.channelId;
            x['customerConversationId'] = conversation.customer.conversation.id;
            if (conversation.agent) {
                x['agentId'] = conversation.agent.user.id;
                x['agentName'] = conversation.agent.user.name;
                x['agentChannelId'] = conversation.agent.channelId;
                x['agentConversationId'] = conversation.agent.conversation.id;
            }
            indexExports._appInsights.client.trackEvent("Transcript", x);    
        }

        return await this.updateConversation(conversation);
    }
    async updateLeadConversation(by: By, message: builder.IMessage, from: string): Promise<boolean> {
        
        
       //find the latest converation by customer id, channel, bot
        let  conversations = await ConversationModel.find({ 'customer.user.id': by.customerId });
        conversations=    conversations.filter(conversation => conversation.customer.channelId === message.address.channelId && conversation.customer.bot.name === message.address.bot.name && conversation.transcript.length>0 );               
        conversations= conversations.sort((x, y) => y.transcript[y.transcript.length - 1].timestamp - x.transcript[x.transcript.length - 1].timestamp);
        
        let conversation = conversations.length>0 && conversations[conversations.length-1] ;

               
        if (!conversation) return false;
        //Check if Lead exists - if not create
        let  lead = await LeadModel.findOne({'leadId': by.customerId });
        if(!lead){
        //create lead and add converation 
        lead = await  this.createLead(conversation.customer.user.id,conversation.customer.user.name)
        }

        await this.updateLead(lead,conversation);
       
    }


    async connectCustomerToAgent(by: By, agentAddress: builder.IAddress): Promise<Conversation> {
        const conversation: Conversation = await this.getConversation(by);
        if (conversation) {
            conversation.state = ConversationState.Agent;
            conversation.agent = agentAddress;
        }
        const success = await this.updateConversation(conversation);
        if (success)
            return conversation;
        else
            return null;
    }

    async queueCustomerForAgent(by: By): Promise<boolean> {
        const conversation: Conversation = await this.getConversation(by);
        if (!conversation) {
            return false;
        } else {
            conversation.state = ConversationState.Waiting;
            return await this.updateConversation(conversation);
        }
    }

    async connectCustomerToBot(by: By): Promise<boolean> {
        const conversation: Conversation = await this.getConversation(by);
        if (!conversation) {
            return false;
        } else {
            conversation.state = ConversationState.Bot;
            if (indexExports._retainData === "true") {
                //if retain data is true, AND the user has spoken to an agent - delete the agent record  
                //this is necessary to avoid a bug where the agent cannot connect to another user after disconnecting with a user
                if (conversation.agent) {
                    conversation.agent = null;
                    return await this.updateConversation(conversation);
                } else {
                    //otherwise, just update the conversation
                    return await this.updateConversation(conversation);
                }
            } else {
                //if retain data is false, delete the whole conversation after talking to agent
                if (conversation.agent) {
                    return await this.deleteConversation(conversation);
                } else {
                    //otherwise, just update the conversation
                    return await this.updateConversation(conversation);
                }
            }
        }
    }

    async getConversation(by: By, customerAddress?: builder.IAddress): Promise<Conversation> {
        if (by.customerName) {
            const conversation = await ConversationModel.findOne({ 'customer.user.name': by.customerName });
            return conversation;
        } else if (by.customerId) {
            const conversation = await ConversationModel.findOne({ 'customer.user.id': by.customerId });
            return conversation;
        }  else if (by.agentConversationId) {
            const conversation = await ConversationModel.findOne({ 'agent.conversation.id': by.agentConversationId });
            if (conversation) return conversation;
            else return null;
        } else if (by.customerConversationId) {
            let conversation: Conversation = await ConversationModel.findOne({ 'customer.conversation.id': by.customerConversationId });
            if (!conversation && customerAddress) {
             conversation = await this.createConversation(customerAddress);                   
             return conversation;    
            }
            else
            {
            return conversation;
            }
            
        } else if (by.bestChoice){
            const waitingLongest = (await this.getCurrentConversations())
                .filter(conversation => conversation.state === ConversationState.Waiting)
                .sort((x, y) => y.transcript[y.transcript.length - 1].timestamp - x.transcript[x.transcript.length - 1].timestamp);
            return waitingLongest.length > 0 && waitingLongest[0];
        }
        return null;
    }

    async getCurrentConversations(): Promise<Conversation[]> {
        let conversations;
        try {
            conversations = await ConversationModel.find();
        } catch (error) {
            console.log('Failed loading conversations');
            console.log(error);
        }
        return conversations;
    }

    private async createConversation(customerAddress: builder.IAddress): Promise<Conversation> {
        // var obj = {
        //     customer: customerAddress,
        //     state: ConversationState.Bot,
        //     transcript: []
        // };
        // var id = customerAddress.conversation.id;
        // return new Promise<Conversation>((resolve, reject)=>{
        //     ConversationModel.update({'customer.conversation.id':id}, obj, { upsert: true }).then( async (conv)  =>  {
        //         console.log('promise handled')
        //         var conversation = await ConversationModel.findOne({ 'customer.conversation.id': id });
        //         return conversation;
                               
        //     }).catch(async (error)=>{
        //         console.log('promise not handled handled')
        //         var conversation = await ConversationModel.findOne({ 'customer.conversation.id': id });
        //         return conversation;
        //     });        
        // });
          return await ConversationModel.create({
            customer: customerAddress,
            state: ConversationState.Bot,
            transcript: []
        });
    }

    private async updateConversation(conversation: Conversation): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            ConversationModel.findByIdAndUpdate((conversation as any)._id, conversation).then((error) => {
                resolve(true)
            }).catch((error) => {
                console.log('Failed to update conversation');
                console.log(conversation as any);
                resolve(false);
            });
        });
    }

    private async deleteConversation(conversation: Conversation): Promise<boolean> {
        return new Promise<boolean>((resolve) => {
            ConversationModel.findByIdAndRemove((conversation as any)._id).then((error) => {
                resolve(true);
            })
        });
    }

    private async createLead(id:string, name:string): Promise<Lead> {
        return await LeadModel.create({
            leadId: id,
            name: name           
        });
    }

    private async updateLead(lead: Lead,conv:Conversation): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            if (!lead.lastConversationsByChannel || lead.lastConversationsByChannel.length<=0){
                lead.lastConversationsByChannel= [conv];
            }else{
              let convs= lead.lastConversationsByChannel.filter(conversation => conversation.customer.channelId === conv.customer.channelId && conversation.customer.bot.name === conv.customer.bot.name);
              if(!convs || convs.length<=0){
                  lead.lastConversationsByChannel.push(conv);
              }
              else {
                  var index =  lead.lastConversationsByChannel.indexOf(convs[0]);
                  if (index > -1) {
                    lead.lastConversationsByChannel.splice(index, 1);
                  }
                  lead.lastConversationsByChannel.push(conv);
              }
            }
            LeadModel.findByIdAndUpdate((lead as any)._id, {lastConversationsByChannel:lead.lastConversationsByChannel}).then((error) => {
                resolve(true)
            }).catch((error) => {
                console.log('Failed to update lead');
                console.log(lead as any);
                resolve(false);
            });
        });
    }

    private async deleteLead(lead: Lead): Promise<boolean> {
        return new Promise<boolean>((resolve) => {
            LeadModel.findByIdAndRemove((lead as any)._id).then((error) => {
                resolve(true);
            })
        });
    }

    private async collectSentiment(text: string): Promise<number> {
        if (text == null || text == '') return;
        let _sentimentUrl = 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment';
        let _sentimentId = 'bot-analytics';
        let _sentimentKey = indexExports._textAnalyticsKey;

        let options = {
            url: _sentimentUrl,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Ocp-Apim-Subscription-Key': _sentimentKey
            },
            json: true,
            body: {
                "documents": [
                    {
                        "language": "en",
                        "id": _sentimentId,
                        "text": text
                    }
                ]
            }
        };

        return new Promise<number>(function (resolve, reject) {
            request(options, (error, response, body) => {
                if (error) { reject(error); }
                let result: any = _.find(body.documents, { id: _sentimentId }) || {};
                let score = result.score || null;
                resolve(score);
            });
        });
    }
}
