# -*- coding: utf-8 -*-

from openerp import SUPERUSER_ID
from openerp.osv import osv, orm, fields
from openerp.addons.web import http
from openerp.addons.web.http import request
import openerp.addons.website_quote.controllers.main as main
import werkzeug
import datetime
import time

class sale_order(osv.Model):
    _inherit = "sale.order"
    _columns = {
        'customize_image': fields.binary("Customized Image",
            help="This field holds the image used for customize your product, limited to 1024x1024px"),
        'customize_ai': fields.binary("Customized AI",
            help="This field holds the AI file (photoshop layout)", readonly=True),
        'customize_ai_name': fields.char('AI File name', 40, readonly=True),
        'attachment_ids': fields.many2many('ir.attachment', 'sale_order_ir_attachments_rel', 'sale_order_id', 'attachment_id', 'Attachments'),
        'state': fields.selection([
            ('draft', 'Draft Quotation'),
            ('sent', 'Quotation Sent'),
            ('cancel', 'Cancelled'),
            ('justify_image', 'Wait for AI layout generated'),
            ('image_rejected', 'Image has been rejected'),
            ('justify_ai', 'Wait for customer accept AI layout'),
            ('ai_rejected', 'AI layout rejected'),
            ('waiting_date', 'Waiting Schedule'),
            ('progress', 'Sales Order'),
            ('manual', 'Sale to Invoice'),
            ('shipping_except', 'Shipping Exception'),
            ('invoice_except', 'Invoice Exception'),
            ('done', 'Done'),            
            ], 'Status', readonly=True, copy=False, help="Gives the status of the quotation or sales order.\
              \nThe exception status is automatically set when a cancel operation occurs \
              in the invoice validation (Invoice Exception) or in the picking list process (Shipping Exception).\nThe 'Waiting Schedule' status is set when the invoice is confirmed\
               but waiting for the scheduler to run on the order date.", select=True),
    }

class res_partner(osv.Model):
    _inherit = 'res.partner'
    _columns = {
        'qblistid': fields.char(),
        'edit_sequence': fields.char(),
        'is_qb_notification': fields.char()
    }

class product_product(osv.Model):
    _inherit = 'product.product'
    _columns = {
        'qblistid': fields.char(),
        'edit_sequence': fields.char()
    }

class product_category(osv.Model):
    _inherit = 'product.category'
    _columns = {
        'qblistid': fields.char(),
        'edit_sequence': fields.char(),
        'is_qb_notification': fields.char()
    }

class product_template(osv.Model):
    _inherit = 'product.template'
    _columns = {
        'lag_prodid': fields.char(),
        'qblistid': fields.char(),
        'edit_sequence': fields.char(),
        'is_qb_notification': fields.char()
    }

class sale_order(osv.osv):
    _inherit = 'sale.order'
    _description = 'customize '
    def action_button_confirm(self, cr, uid, ids, context=None):
        self.write(cr, uid, ids, {'state': 'justify_image'})
        return super(sale_order, self).action_button_confirm(cr, uid, ids, context)
    def approve_image(self, cr, uid, ids, context=None):
        '''
        This function prints the sales order and mark it as sent, so that we can see more easily the next step of the workflow
        '''
        assert len(ids) == 1, 'This option should only be used for a single id at a time'
        self.write(cr, uid, ids, {'state': 'justify_ai'})
        self.signal_workflow(cr, uid, ids, 'image_approve')
        return True

    def disapprove_image(self, cr, uid, ids, context=None):
        '''
        This function prints the sales order and mark it as sent, so that we can see more easily the next step of the workflow
        '''
        assert len(ids) == 1, 'This option should only be used for a single id at a time'
        self.write(cr, uid, ids, {'state': 'image_rejected'})
        self.signal_workflow(cr, uid, ids, 'image_disapprove')
        return True

class sale_quote_extension(main.sale_quote):
    @http.route()
    def accept(self, order_id, token=None, signer=None, sign=None, **post):
        order_obj = request.registry.get('sale.order')
        order = order_obj.browse(request.cr, SUPERUSER_ID, order_id)
        order_obj.write(request.cr, SUPERUSER_ID,  [order_id], {'state': 'justify_image'}, context=request.context)
        return super(sale_quote_extension, self).accept(order_id, token, signer, sign)

    @http.route(['/quote/reaccept/<int:order_id>/<token>'], type='http', auth="public", website=True)
    def reaccept(self, order_id, token=None, signer=None, sign=None, **post):
        order_obj = request.registry.get('sale.order')
        order = order_obj.browse(request.cr, SUPERUSER_ID, order_id)
        order_obj.write(request.cr, SUPERUSER_ID,  [order_id], {'state': 'justify_image'}, context=request.context)
        order_obj.signal_workflow(request.cr, SUPERUSER_ID, [order_id], 'order_confirm', context=request.context)       
        return werkzeug.utils.redirect("/quote/%s/%s?message=5" % (order_id, token))

    @http.route(['/quote/recancel/<int:order_id>/<token>'], type='http', auth="public", website=True)
    def recancel(self, order_id, token=None, signer=None, sign=None, **post):
        order_obj = request.registry.get('sale.order')
        order = order_obj.browse(request.cr, SUPERUSER_ID, order_id)
        order_obj.write(request.cr, SUPERUSER_ID,  [order_id], {'state': 'justify_image'}, context=request.context)
        order_obj.signal_workflow(request.cr, SUPERUSER_ID, [order_id], 'order_confirm', context=request.context)       
        return werkzeug.utils.redirect("/quote/%s/%s?message=2" % (order_id, token))

    @http.route(['/quote/acceptai/<int:order_id>/<token>'], type='http', auth="public", website=True)
    def acceptai(self, order_id, token=None, signer=None, sign=None, **post):
        order_obj = request.registry.get('sale.order')
        order = order_obj.browse(request.cr, SUPERUSER_ID, order_id)
        order_obj.write(request.cr, SUPERUSER_ID,  [order_id], {'state': 'manual'}, context=request.context)
        order_obj.signal_workflow(request.cr, SUPERUSER_ID, [order_id], 'ai_approve', context=request.context)       
        return werkzeug.utils.redirect("/quote/%s/%s?message=4" % (order_id, token))

    @http.route(['/quote/rejectai/<int:order_id>/<token>'], type='http', auth="public", website=True)
    def rejectai(self, order_id, token=None, signer=None, sign=None, **post):
        order_obj = request.registry.get('sale.order')
        order = order_obj.browse(request.cr, SUPERUSER_ID, order_id)
        order_obj.write(request.cr, SUPERUSER_ID,  [order_id], {'state': 'ai_rejected'}, context=request.context)
        order_obj.signal_workflow(request.cr, SUPERUSER_ID, [order_id], 'ai_disapprove', context=request.context)       
        return werkzeug.utils.redirect("/quote/%s/%s?message=6" % (order_id, token))

# class carryout_customization(models.Model):
#     _name = 'carryout_customization.carryout_customization'

#     name = fields.Char()