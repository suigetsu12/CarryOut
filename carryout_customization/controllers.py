# -*- coding: utf-8 -*-
from openerp import http

# class CarryoutCustomization(http.Controller):
#     @http.route('/carryout_customization/carryout_customization/', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/carryout_customization/carryout_customization/objects/', auth='public')
#     def list(self, **kw):
#         return http.request.render('carryout_customization.listing', {
#             'root': '/carryout_customization/carryout_customization',
#             'objects': http.request.env['carryout_customization.carryout_customization'].search([]),
#         })

#     @http.route('/carryout_customization/carryout_customization/objects/<model("carryout_customization.carryout_customization"):obj>/', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('carryout_customization.object', {
#             'object': obj
#         })